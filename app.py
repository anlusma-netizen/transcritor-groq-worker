import os
import re
import json
import math
import shutil
import tempfile
import subprocess
from pathlib import Path
from typing import Optional, List, Dict, Any
from urllib.parse import urlparse, parse_qs

import requests
from fastapi import FastAPI, UploadFile, File, Request, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from docx import Document
from docx.shared import Pt
from groq import Groq

app = FastAPI(title="Worker Telegram → Groq → DOCX", version="4.0.0")

GROQ_API_KEY = os.getenv("GROQ_API_KEY", "")
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
GROQ_TRANSCRIPTION_MODEL = os.getenv("GROQ_TRANSCRIPTION_MODEL", "whisper-large-v3-turbo")
GROQ_TRANSLATION_MODEL = os.getenv("GROQ_TRANSLATION_MODEL", "llama-3.3-70b-versatile")
MAX_AUDIO_MB = int(os.getenv("MAX_AUDIO_MB", "24"))
TARGET_AUDIO_BITRATE = os.getenv("TARGET_AUDIO_BITRATE", "24k")
TARGET_AUDIO_FORMAT = os.getenv("TARGET_AUDIO_FORMAT", "mp3")

client = Groq(api_key=GROQ_API_KEY) if GROQ_API_KEY else None


@app.get("/")
def root():
    return {
        "ok": True,
        "service": "transcritor-groq-worker",
        "version": "4.0.0",
        "routes": [
            "/health",
            "/process-telegram-media",
            "/process-source",
        ],
    }


@app.get("/health")
def health():
    return {
        "ok": True,
        "version": "4.0.0",
        "groq_key_configured": bool(GROQ_API_KEY),
        "telegram_token_configured": bool(TELEGRAM_BOT_TOKEN),
        "transcription_model": GROQ_TRANSCRIPTION_MODEL,
        "translation_model": GROQ_TRANSLATION_MODEL,
    }


def run_cmd(cmd: List[str]):
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"Erro no comando: {' '.join(cmd)}\nSTDERR:\n{result.stderr}")
    return result


def safe_filename(name: str, fallback: str = "arquivo") -> str:
    name = name or fallback
    name = re.sub(r"[^\w\-. ]+", "_", name, flags=re.UNICODE).strip()
    return name[:120] or fallback


def download_url(url: str, output_path: Path):
    with requests.get(url, stream=True, timeout=120) as r:
        r.raise_for_status()
        with open(output_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    f.write(chunk)


def google_drive_direct_url(url: str) -> str:
    # Suporta links comuns do Drive:
    # https://drive.google.com/file/d/ID/view?usp=sharing
    # https://drive.google.com/open?id=ID
    # https://drive.google.com/uc?id=ID
    if "drive.google.com" not in url:
        return url

    file_id = None
    m = re.search(r"/file/d/([^/]+)", url)
    if m:
        file_id = m.group(1)
    else:
        qs = parse_qs(urlparse(url).query)
        if "id" in qs:
            file_id = qs["id"][0]

    if not file_id:
        return url

    return f"https://drive.google.com/uc?export=download&id={file_id}"


def download_google_drive_or_url(url: str, output_path: Path):
    url = google_drive_direct_url(url)

    # Tentativa simples primeiro.
    session = requests.Session()
    response = session.get(url, stream=True, timeout=120)

    # Google Drive pode exigir confirm token para arquivo grande.
    token = None
    for key, value in response.cookies.items():
        if key.startswith("download_warning"):
            token = value
            break

    if token:
        parsed = urlparse(url)
        qs = parse_qs(parsed.query)
        file_id = qs.get("id", [None])[0]
        if file_id:
            url = f"https://drive.google.com/uc?export=download&confirm={token}&id={file_id}"
            response = session.get(url, stream=True, timeout=120)

    response.raise_for_status()
    with open(output_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=1024 * 1024):
            if chunk:
                f.write(chunk)


def download_telegram_file(file_id: str, output_path: Path):
    if not TELEGRAM_BOT_TOKEN:
        raise RuntimeError("TELEGRAM_BOT_TOKEN não configurado no Railway.")

    info_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getFile"
    info = requests.get(info_url, params={"file_id": file_id}, timeout=60).json()

    if not info.get("ok"):
        raise RuntimeError(f"Telegram getFile falhou: {info}")

    file_path = info["result"]["file_path"]
    file_url = f"https://api.telegram.org/file/bot{TELEGRAM_BOT_TOKEN}/{file_path}"
    download_url(file_url, output_path)


def convert_to_audio(input_path: Path, output_path: Path):
    cmd = [
        "ffmpeg", "-y",
        "-i", str(input_path),
        "-vn",
        "-ac", "1",
        "-ar", "16000",
        "-b:a", TARGET_AUDIO_BITRATE,
        str(output_path),
    ]
    run_cmd(cmd)


def file_mb(path: Path) -> float:
    return path.stat().st_size / (1024 * 1024)


def get_duration_seconds(path: Path) -> float:
    cmd = [
        "ffprobe", "-v", "error",
        "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1",
        str(path)
    ]
    result = run_cmd(cmd)
    try:
        return float(result.stdout.strip())
    except Exception:
        return 0.0


def split_audio(input_audio: Path, chunks_dir: Path, chunk_seconds: int = 600) -> List[Path]:
    chunks_dir.mkdir(parents=True, exist_ok=True)
    pattern = chunks_dir / f"chunk_%03d.{TARGET_AUDIO_FORMAT}"
    cmd = [
        "ffmpeg", "-y",
        "-i", str(input_audio),
        "-f", "segment",
        "-segment_time", str(chunk_seconds),
        "-c", "copy",
        str(pattern)
    ]
    run_cmd(cmd)
    return sorted(chunks_dir.glob(f"chunk_*.{TARGET_AUDIO_FORMAT}"))


def transcribe_one(audio_path: Path) -> Dict[str, Any]:
    if not client:
        raise RuntimeError("GROQ_API_KEY não configurada no Railway.")

    with open(audio_path, "rb") as f:
        result = client.audio.transcriptions.create(
            file=(audio_path.name, f.read()),
            model=GROQ_TRANSCRIPTION_MODEL,
            response_format="verbose_json",
            temperature=0,
        )

    if hasattr(result, "model_dump"):
        return result.model_dump()
    if isinstance(result, dict):
        return result
    return json.loads(result.json())


def transcribe_audio(audio_path: Path, workdir: Path) -> Dict[str, Any]:
    max_mb = MAX_AUDIO_MB
    if file_mb(audio_path) <= max_mb:
        return transcribe_one(audio_path)

    chunks = split_audio(audio_path, workdir / "chunks", chunk_seconds=600)
    all_text = []
    all_segments = []
    language = None
    offset = 0.0

    for chunk in chunks:
        chunk_duration = get_duration_seconds(chunk)
        res = transcribe_one(chunk)
        if not language:
            language = res.get("language")
        text = res.get("text", "")
        if text:
            all_text.append(text)

        for seg in res.get("segments", []) or []:
            seg = dict(seg)
            if "start" in seg:
                seg["start"] = float(seg["start"]) + offset
            if "end" in seg:
                seg["end"] = float(seg["end"]) + offset
            all_segments.append(seg)

        offset += chunk_duration

    return {
        "text": "\n".join(all_text).strip(),
        "language": language,
        "segments": all_segments,
    }


def translate_to_ptbr(text: str) -> str:
    if not text.strip():
        return ""

    if not client:
        raise RuntimeError("GROQ_API_KEY não configurada no Railway.")

    # Divide para evitar prompt grande demais.
    parts = []
    max_chars = 9000
    current = []
    current_len = 0
    for paragraph in text.split("\n"):
        if current_len + len(paragraph) > max_chars and current:
            parts.append("\n".join(current))
            current = []
            current_len = 0
        current.append(paragraph)
        current_len += len(paragraph) + 1
    if current:
        parts.append("\n".join(current))

    translated_parts = []
    for i, part in enumerate(parts, start=1):
        prompt = f"""
Traduza o texto abaixo para português brasileiro.

Regras:
- Mantenha máxima fidelidade ao texto original.
- Não resuma.
- Não melhore a copy.
- Não adicione argumentos.
- Preserve repetições, promessas, ganchos, CTAs e estrutura de VSL/anúncio.
- Se houver frases quebradas, mantenha natural em PT-BR sem inventar conteúdo.
- Entregue apenas a tradução.

Texto:
{part}
""".strip()

        completion = client.chat.completions.create(
            model=GROQ_TRANSLATION_MODEL,
            messages=[
                {"role": "system", "content": "Você é um tradutor profissional de copy, VSL e anúncios para português brasileiro."},
                {"role": "user", "content": prompt},
            ],
            temperature=0,
        )
        translated_parts.append(completion.choices[0].message.content.strip())

    return "\n\n".join(translated_parts).strip()


def fmt_time(seconds: float) -> str:
    seconds = int(seconds or 0)
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def create_docx(original_name: str, transcription: Dict[str, Any], translated_text: str, output_path: Path):
    doc = Document()

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(11)

    title = doc.add_heading("Transcrição / Tradução PT-BR", level=1)
    doc.add_paragraph(f"Arquivo: {original_name}")
    doc.add_paragraph(f"Idioma detectado: {transcription.get('language') or 'não identificado'}")
    doc.add_paragraph("Formato: Copy/VSL — blocos fáceis de ler")
    doc.add_paragraph("")

    doc.add_heading("Versão em português brasileiro", level=2)
    for block in translated_text.split("\n\n"):
        if block.strip():
            doc.add_paragraph(block.strip())

    doc.add_page_break()
    doc.add_heading("Transcrição original", level=2)

    segments = transcription.get("segments") or []
    if segments:
        for seg in segments:
            start = fmt_time(seg.get("start", 0))
            end = fmt_time(seg.get("end", 0))
            text = (seg.get("text") or "").strip()
            if text:
                p = doc.add_paragraph()
                p.add_run(f"[{start} – {end}] ").bold = True
                p.add_run(text)
    else:
        doc.add_paragraph(transcription.get("text", ""))

    doc.save(output_path)


def process_file(input_path: Path, original_name: str, workdir: Path) -> Path:
    audio_path = workdir / f"audio_convertido.{TARGET_AUDIO_FORMAT}"
    convert_to_audio(input_path, audio_path)
    transcription = transcribe_audio(audio_path, workdir)
    translated = translate_to_ptbr(transcription.get("text", ""))
    out_name = safe_filename(Path(original_name).stem or "transcricao") + "_ptbr.docx"
    output_path = workdir / out_name
    create_docx(original_name, transcription, translated, output_path)
    return output_path


@app.post("/process-telegram-media")
async def process_telegram_media(request: Request, file: UploadFile = File(None)):
    """
    Endpoint antigo: recebe o binário diretamente do n8n.
    Mantido para áudio pequeno que já estava funcionando.
    """
    if file is None:
        # Alguns HTTP Request do n8n mandam binário cru sem multipart.
        body = await request.body()
        if not body:
            raise HTTPException(status_code=400, detail="Nenhum arquivo recebido.")
        original_name = request.headers.get("x-file-name", "arquivo_telegram")
        with tempfile.TemporaryDirectory() as tmp:
            workdir = Path(tmp)
            input_path = workdir / safe_filename(original_name, "input.bin")
            input_path.write_bytes(body)
            output_path = process_file(input_path, original_name, workdir)
            final_path = Path(tempfile.gettempdir()) / output_path.name
            shutil.copy(output_path, final_path)
            return FileResponse(final_path, filename=output_path.name, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    original_name = request.headers.get("x-file-name", file.filename or "arquivo_telegram")
    with tempfile.TemporaryDirectory() as tmp:
        workdir = Path(tmp)
        input_path = workdir / safe_filename(original_name, "input.bin")
        with open(input_path, "wb") as f:
            f.write(await file.read())
        output_path = process_file(input_path, original_name, workdir)
        final_path = Path(tempfile.gettempdir()) / output_path.name
        shutil.copy(output_path, final_path)
        return FileResponse(final_path, filename=output_path.name, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.post("/process-source")
async def process_source(payload: Dict[str, Any]):
    """
    Endpoint novo:
    - sourceType=url: baixa arquivo por link público
    - sourceType=telegram_file_id: baixa direto do Telegram usando TELEGRAM_BOT_TOKEN
    """
    source_type = payload.get("sourceType")
    original_name = payload.get("fileName") or "arquivo"

    with tempfile.TemporaryDirectory() as tmp:
        workdir = Path(tmp)
        input_path = workdir / safe_filename(original_name, "input.bin")

        if source_type == "url":
            url = payload.get("url")
            if not url:
                raise HTTPException(status_code=400, detail="URL não enviada.")
            download_google_drive_or_url(url, input_path)

        elif source_type == "telegram_file_id":
            file_id = payload.get("fileId")
            if not file_id:
                raise HTTPException(status_code=400, detail="fileId do Telegram não enviado.")
            download_telegram_file(file_id, input_path)

        else:
            raise HTTPException(status_code=400, detail="Envie arquivo ou link público.")

        output_path = process_file(input_path, original_name, workdir)
        final_path = Path(tempfile.gettempdir()) / output_path.name
        shutil.copy(output_path, final_path)
        return FileResponse(final_path, filename=output_path.name, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

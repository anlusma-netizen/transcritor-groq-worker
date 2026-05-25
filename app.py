import os
import re
import json
import math
import shutil
import tempfile
import time
import subprocess
from pathlib import Path
from typing import Optional, List, Dict, Any
from urllib.parse import urlparse, parse_qs
from html import escape

import requests
import gdown
from fastapi import FastAPI, UploadFile, File, Request, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from docx import Document
from docx.shared import Pt
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.enums import TA_CENTER
from groq import Groq

app = FastAPI(title="Worker Telegram → Groq → DOCX", version="9.0.0")

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    return JSONResponse(
        status_code=500,
        content={
            "ok": False,
            "error": str(exc),
            "type": exc.__class__.__name__,
            "hint": "Veja os logs do Railway para detalhes completos."
        },
    )


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
        "version": "9.0.0",
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
        "version": "9.0.0",
        "groq_key_configured": bool(GROQ_API_KEY),
        "telegram_token_configured": bool(TELEGRAM_BOT_TOKEN),
        "transcription_model": GROQ_TRANSCRIPTION_MODEL,
        "translation_model": GROQ_TRANSLATION_MODEL,
        "translation_chunk_chars": os.getenv("TRANSLATION_CHUNK_CHARS", "3000"),
        "translation_delay_seconds": os.getenv("TRANSLATION_DELAY_SECONDS", "12"),
        "output_format": "pdf",
        "include_original_for_non_pt": os.getenv("INCLUDE_ORIGINAL_FOR_NON_PT", "true"),
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


def extract_google_drive_file_id(url: str) -> Optional[str]:
    if "drive.google.com" not in url:
        return None

    m = re.search(r"/file/d/([^/]+)", url)
    if m:
        return m.group(1)

    qs = parse_qs(urlparse(url).query)
    if "id" in qs:
        return qs["id"][0]

    return None


def download_google_drive_or_url(url: str, output_path: Path):
    """
    Baixa arquivo por URL.
    Para Google Drive, usa gdown, que lida melhor com páginas de confirmação
    de arquivos grandes.
    """
    file_id = extract_google_drive_file_id(url)

    if file_id:
        gdown_url = f"https://drive.google.com/uc?id={file_id}"
        result = gdown.download(gdown_url, str(output_path), quiet=False, fuzzy=True)
        if not result or not output_path.exists() or output_path.stat().st_size < 1024:
            raise RuntimeError(
                "Não consegui baixar o arquivo do Google Drive. "
                "Confirme que o compartilhamento está como 'Qualquer pessoa com o link pode visualizar'."
            )
        return

    # URL comum fora do Drive
    download_url(url, output_path)

    if not output_path.exists() or output_path.stat().st_size < 1024:
        raise RuntimeError("O download do link falhou ou retornou arquivo vazio.")


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




def is_portuguese_language(language: Optional[str]) -> bool:
    if not language:
        return False
    lang = str(language).lower().strip()
    return lang in {"pt", "por", "portuguese", "português", "portugues", "pt-br", "pt_br"} or "portugu" in lang


def split_text_for_translation(text: str, max_chars: int = 3000) -> List[str]:
    """
    Divide o texto em blocos pequenos para não estourar o limite de tokens por minuto da Groq.
    """
    raw_paragraphs = text.split("\n")
    paragraphs = [p.strip() for p in raw_paragraphs if p.strip()]

    parts = []
    current = []
    current_len = 0

    for paragraph in paragraphs:
        if len(paragraph) > max_chars:
            chunks = [paragraph[i:i + max_chars] for i in range(0, len(paragraph), max_chars)]
        else:
            chunks = [paragraph]

        for chunk in chunks:
            if current and current_len + len(chunk) + 1 > max_chars:
                parts.append("\n".join(current).strip())
                current = []
                current_len = 0

            current.append(chunk)
            current_len += len(chunk) + 1

    if current:
        parts.append("\n".join(current).strip())

    return [p for p in parts if p]


def run_chat_with_retry(prompt: str, index: int, total: int) -> str:
    last_error = None

    for attempt in range(1, 6):
        try:
            completion = client.chat.completions.create(
                model=GROQ_TRANSLATION_MODEL,
                messages=[
                    {
                        "role": "system",
                        "content": "Você é um especialista em tradução fiel, organização de copy, VSL e cartas de vendas em português brasileiro."
                    },
                    {"role": "user", "content": prompt},
                ],
                temperature=0,
            )
            return completion.choices[0].message.content.strip()

        except Exception as exc:
            last_error = exc
            error_text = str(exc)

            if (
                "rate_limit_exceeded" in error_text
                or "Request too large" in error_text
                or "tokens per minute" in error_text
                or "TPM" in error_text
                or "Rate limit" in error_text
            ):
                wait_seconds = int(os.getenv("TRANSLATION_RETRY_WAIT_SECONDS", "70"))
                print(
                    f"Rate limit no bloco {index}/{total}. "
                    f"Tentativa {attempt}/5. Aguardando {wait_seconds}s."
                )
                time.sleep(wait_seconds)
                continue

            raise

    raise RuntimeError(f"Falha ao processar bloco {index}/{total} após retries: {last_error}")


def process_text_to_ptbr_for_pdf(text: str, source_is_portuguese: bool) -> str:
    """
    Retorna o texto final em PT-BR, limpo para PDF.
    """
    if not text.strip():
        return ""

    if not client:
        raise RuntimeError("GROQ_API_KEY não configurada no Railway.")

    chunk_chars = int(os.getenv("TRANSLATION_CHUNK_CHARS", "3000"))
    delay_seconds = int(os.getenv("TRANSLATION_DELAY_SECONDS", "12"))

    parts = split_text_for_translation(text, max_chars=chunk_chars)
    processed_parts = []

    for i, part in enumerate(parts, start=1):
        if source_is_portuguese:
            instruction = f"""
Organize o bloco {i}/{len(parts)} abaixo em português brasileiro para virar um PDF de estudo de copy/VSL/carta de vendas.

Regras:
- O texto já está em português. Não traduza.
- Não resuma.
- Não reescreva criativamente.
- Preserve o máximo possível as palavras originais.
- Apenas limpe pontuação quando necessário e quebre em parágrafos fáceis de ler.
- Não inclua timestamps.
- Não repita o texto.
- Use **negrito** somente em poucas frases realmente importantes: promessa central, mecanismo, prova, dor principal, grande objeção, oferta e chamada para ação.
- Não coloque negrito em tudo.
- Entregue apenas o texto final.

Texto:
{part}
""".strip()
        else:
            instruction = f"""
Traduza o bloco {i}/{len(parts)} abaixo para português brasileiro e organize para virar um PDF de estudo de copy/VSL/carta de vendas.

Regras:
- Mantenha máxima fidelidade ao texto original.
- Não resuma.
- Não melhore a copy.
- Não adicione argumentos.
- Preserve repetições, promessas, ganchos, CTAs e estrutura de VSL/anúncio.
- Não inclua timestamps.
- Use **negrito** somente em poucas frases realmente importantes: promessa central, mecanismo, prova, dor principal, grande objeção, oferta e chamada para ação.
- Não coloque negrito em tudo.
- Entregue apenas o texto final traduzido.

Texto:
{part}
""".strip()

        print(f"Processando texto para PDF bloco {i}/{len(parts)} com {len(part)} caracteres...")
        processed_parts.append(run_chat_with_retry(instruction, i, len(parts)))

        if i < len(parts):
            time.sleep(delay_seconds)

    return "\n\n".join(processed_parts).strip()


def fmt_time(seconds: float) -> str:

    seconds = int(seconds or 0)
    h = seconds // 3600
    m = (seconds % 3600) // 60
    s = seconds % 60
    return f"{h:02d}:{m:02d}:{s:02d}"



def markdown_bold_to_reportlab(text: str) -> str:
    """
    Converte **negrito** em tags <b> seguras para Paragraph do ReportLab.
    """
    text = escape(text)
    parts = text.split("**")
    out = []
    bold = False
    for part in parts:
        if bold:
            out.append(f"<b>{part}</b>")
        else:
            out.append(part)
        bold = not bold
    return "".join(out).replace("\n", "<br/>")


def make_pdf_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    page_text = f"Página {doc.page}"
    canvas.drawRightString(A4[0] - 1.8 * cm, 1.2 * cm, page_text)
    canvas.restoreState()


def create_pdf(
    original_name: str,
    transcription: Dict[str, Any],
    final_text: str,
    output_path: Path,
    include_original: bool = False,
):
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        title="Transcrição e análise de copy",
        author="Transcritor Groq",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleCustom",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        alignment=TA_CENTER,
        spaceAfter=18,
    )
    h2_style = ParagraphStyle(
        "H2Custom",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        spaceBefore=14,
        spaceAfter=8,
    )
    body_style = ParagraphStyle(
        "BodyCustom",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10.5,
        leading=15,
        spaceAfter=8,
    )
    meta_style = ParagraphStyle(
        "MetaCustom",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        spaceAfter=4,
    )

    story = []
    story.append(Paragraph("Transcrição em PT-BR - Copy/VSL", title_style))
    story.append(Paragraph(f"<b>Arquivo:</b> {escape(original_name)}", meta_style))
    story.append(Paragraph(f"<b>Idioma detectado:</b> {escape(str(transcription.get('language') or 'não identificado'))}", meta_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Texto final", h2_style))

    paragraphs = [p.strip() for p in final_text.split("\n\n") if p.strip()]
    if not paragraphs:
        paragraphs = [final_text.strip()] if final_text.strip() else ["Sem texto final disponível."]

    for p in paragraphs:
        story.append(Paragraph(markdown_bold_to_reportlab(p), body_style))

    if include_original:
        story.append(PageBreak())
        story.append(Paragraph("Transcrição original", h2_style))
        original_text = transcription.get("text", "") or ""
        for p in [x.strip() for x in original_text.split("\n\n") if x.strip()]:
            story.append(Paragraph(markdown_bold_to_reportlab(p), body_style))

    doc.build(story, onFirstPage=make_pdf_footer, onLaterPages=make_pdf_footer)


def process_file(input_path: Path, original_name: str, workdir: Path) -> Path:
    audio_path = workdir / f"audio_convertido.{TARGET_AUDIO_FORMAT}"
    convert_to_audio(input_path, audio_path)
    transcription = transcribe_audio(audio_path, workdir)

    detected_language = transcription.get("language")
    source_is_portuguese = is_portuguese_language(detected_language)
    final_text = process_text_to_ptbr_for_pdf(
        transcription.get("text", ""),
        source_is_portuguese=source_is_portuguese,
    )

    include_original_for_non_pt = os.getenv("INCLUDE_ORIGINAL_FOR_NON_PT", "true").lower() in {"1", "true", "yes", "sim"}
    include_original = include_original_for_non_pt and not source_is_portuguese

    out_name = safe_filename(Path(original_name).stem or "transcricao") + "_ptbr.pdf"
    output_path = workdir / out_name
    create_pdf(original_name, transcription, final_text, output_path, include_original=include_original)
    return output_path



@app.post("/process-telegram-media")
async def process_telegram_media(request: Request, file: Optional[UploadFile] = File(default=None)):
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
            return FileResponse(final_path, filename=output_path.name, media_type="application/pdf")

    original_name = request.headers.get("x-file-name", file.filename or "arquivo_telegram")
    with tempfile.TemporaryDirectory() as tmp:
        workdir = Path(tmp)
        input_path = workdir / safe_filename(original_name, "input.bin")
        with open(input_path, "wb") as f:
            f.write(await file.read())
        output_path = process_file(input_path, original_name, workdir)
        final_path = Path(tempfile.gettempdir()) / output_path.name
        shutil.copy(output_path, final_path)
        return FileResponse(final_path, filename=output_path.name, media_type="application/pdf")


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
        return FileResponse(final_path, filename=output_path.name, media_type="application/pdf")

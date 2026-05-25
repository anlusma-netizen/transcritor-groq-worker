# Worker Telegram → Groq → DOCX

Este worker recebe arquivos pequenos via n8n ou recebe metadados/link para arquivos grandes, converte com FFmpeg, transcreve com Groq Whisper, traduz para português brasileiro com Groq Chat e devolve um DOCX.

## Variáveis obrigatórias

- `GROQ_API_KEY`
- `TELEGRAM_BOT_TOKEN` — necessário para baixar arquivos do Telegram pelo `file_id` sem o n8n carregar o vídeo pesado.

## Variáveis opcionais

- `GROQ_TRANSCRIPTION_MODEL=whisper-large-v3-turbo`
- `GROQ_TRANSLATION_MODEL=llama-3.3-70b-versatile`
- `MAX_AUDIO_MB=24`
- `TARGET_AUDIO_BITRATE=24k`
- `TARGET_AUDIO_FORMAT=mp3`

## Endpoints

- `GET /health`
- `POST /process-telegram-media` — compatível com o fluxo antigo, recebe binário do n8n.
- `POST /process-url-media` — recebe JSON com `media_url` e `file_name`.
- `POST /process-source` — endpoint novo para o n8n mandar `sourceType=url` ou `sourceType=telegram_file_id`.

Para arquivos grandes, prefira mandar um link público de Google Drive/Dropbox no Telegram.

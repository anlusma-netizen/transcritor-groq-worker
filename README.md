# Worker Telegram → Groq → DOCX

Este worker recebe um arquivo enviado pelo n8n, converte/comprime com FFmpeg, transcreve com Groq Whisper, traduz para português brasileiro usando Groq Chat e devolve um arquivo DOCX.

## Variáveis obrigatórias

- `GROQ_API_KEY`

## Variáveis opcionais

- `GROQ_TRANSCRIPTION_MODEL=whisper-large-v3-turbo`
- `GROQ_TRANSLATION_MODEL=llama-3.3-70b-versatile`
- `MAX_AUDIO_MB=24`
- `TARGET_AUDIO_BITRATE=24k`
- `TARGET_AUDIO_FORMAT=mp3`

## Endpoints

- `GET /health`
- `POST /process-telegram-media`

O n8n deve enviar o arquivo como corpo binário da requisição e receber o DOCX como resposta binária.

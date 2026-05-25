# Worker Telegram → Groq → DOCX

Versão 4.

## Rotas

- `GET /health`
- `POST /process-telegram-media` — fluxo antigo, recebe binário do n8n
- `POST /process-source` — fluxo novo, recebe JSON com `sourceType=url` ou `sourceType=telegram_file_id`

## Variáveis obrigatórias no Railway

```txt
GROQ_API_KEY
TELEGRAM_BOT_TOKEN
GROQ_TRANSCRIPTION_MODEL
GROQ_TRANSLATION_MODEL
MAX_AUDIO_MB
TARGET_AUDIO_BITRATE
TARGET_AUDIO_FORMAT
```

## Teste

Abra:

```txt
/health
```

A resposta precisa mostrar:

```json
{
  "ok": true,
  "groq_key_configured": true,
  "telegram_token_configured": true
}
```

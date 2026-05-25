# Worker Telegram → Groq → DOCX

Versão 7. Corrige erro de limite de tokens da Groq na tradução, dividindo o texto em blocos menores com retry.

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


## Variáveis opcionais para arquivos grandes

```txt
TRANSLATION_CHUNK_CHARS=4500
TRANSLATION_DELAY_SECONDS=8
```

Se a Groq ainda reclamar de limite de tokens, reduza `TRANSLATION_CHUNK_CHARS` para 3000 e aumente `TRANSLATION_DELAY_SECONDS` para 15.

# GridPath Proxy Server

Thin Vercel serverless proxy for the GridPath desktop app (Cursor-for-Excel).

Two jobs only:

1. **Authenticate** desktop-app requests against Firebase (the app signs in via Firebase Auth client-side; the proxy verifies the ID token server-side using firebase-admin).
2. **Forward** LLM calls to Anthropic / OpenAI without exposing the API keys to the client.

Plus billing/usage/email endpoints used by the desktop app's settings, subscription, and password-reset flows.

## Endpoint surface

```
api/
├── auth/
│   ├── google.ts                   POST — exchange Firebase ID token for proxy session
│   ├── send-password-reset.ts      POST — trigger Resend password reset email
│   └── send-verification.ts        POST — trigger Resend email verification
├── billing/                        Stripe credits + portal (used for subscription)
├── llm/chat.ts                     POST — forward LLM call (Claude / OpenAI / …)
├── usage/summary.ts                GET  — LLM token usage for current user
└── user/settings.ts                GET/PUT — user preferences
```

## Setup

```bash
cd proxy-server
npm install
cp env.example .env.local
# fill in .env.local — see env.example for required vs optional
npm run dev
```

Server runs at `http://localhost:3000`.

## Required environment variables

See [env.example](./env.example) for the full annotated list. The minimum to boot a working `/api/llm/chat` is:

- `FIREBASE_PROJECT_ID`, `FIREBASE_CLIENT_EMAIL`, `FIREBASE_PRIVATE_KEY` — token verification
- `CLAUDE_API_KEY` — Anthropic
- `JWT_SECRET` — session signing
- `DATABASE_URL`, `NEXT_PUBLIC_SUPABASE_URL`, `SUPABASE_SERVICE_ROLE_KEY` — Postgres for usage / credit accounting

Stripe (`STRIPE_SECRET_KEY`, `STRIPE_WEBHOOK_SECRET`) and Resend (`RESEND_API_KEY`) are required once you turn on billing or transactional email respectively.

`OPENAI_API_KEY`, `GEMINI_API_KEY`, `GROK_API_KEY`, `GOOGLE_VERTEX_API_KEY` are all lazy-initialized — only required if a `/api/llm/chat` call explicitly requests that provider.

## Firebase service account

Firebase console → Project Settings → Service Accounts → Generate New Private Key → download JSON. Map fields:

- `project_id` → `FIREBASE_PROJECT_ID`
- `client_email` → `FIREBASE_CLIENT_EMAIL`
- `private_key` → `FIREBASE_PRIVATE_KEY` (keep the literal `\n` escapes — don't unescape them)

## Deploy to Vercel

```bash
npm i -g vercel       # one-time
vercel login          # one-time
vercel --prod         # from inside proxy-server/
```

Then set the env vars in the Vercel dashboard (Settings → Environment Variables). After the first deploy, update the desktop app's `VITE_PROXY_URL` and `src-tauri/src/engine/llm_providers/proxy.rs` `PROXY_URL` to point at the new Vercel URL.

## /api/llm/chat — request shape

```
POST /api/llm/chat
Authorization: Bearer <firebase-id-token-or-jwt>
{
  "provider": "claude",
  "messages": [
    { "role": "system", "content": "..." },
    { "role": "user", "content": "..." }
  ],
  "maxTokens": 4096,
  "sessionId": "spreadsheet_<id>"
}
```

Response:

```
{
  "success": true,
  "content": "...",
  "usage": { "prompt_tokens": N, "completion_tokens": N, "total_tokens": N },
  "provider": "claude"
}
```

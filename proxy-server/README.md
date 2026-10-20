# GridPath proxy server

The GridPath desktop app's sign-in service, a Vercel serverless function.

The app signs users in with Firebase Auth on the client and sends the
Firebase ID token here. The service verifies it with firebase-admin and
returns a session token (a JWT signed with `JWT_SECRET`) that the app keeps
as its signed-in state.

```
api/auth/google.ts    POST { idToken } → { success, token, user }
lib/firebase.ts       Firebase Admin token verification
lib/jwt.ts            session token signing and verification
lib/middleware.ts     CORS (and token checks for authenticated routes)
```

The app's AI calls do not go through this service: they use the user's own
Claude or ChatGPT subscription directly.

## Setup

```bash
cd proxy-server
npm install
cp env.example .env.local   # the four variables in env.example
npm run dev                 # vercel dev, at http://localhost:3000
```

Deploy with `npm run deploy` (`vercel --prod`).

## License

FSL-1.1-Apache-2.0, like the desktop app; see [LICENSE](LICENSE).

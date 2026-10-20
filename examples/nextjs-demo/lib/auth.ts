// Stand-in for your auth: who is making this request. Replace the body with
// your session check (Auth.js's auth(), Clerk's auth(), a JWT verify…). With
// cookie sessions the browser sends the cookie itself and httpStore() needs
// no `headers`.
const DEMO_TOKENS: Record<string, string> = { "demo-token": "demo-user" };

export async function userFrom(req: Request): Promise<string | null> {
  const token = /^Bearer (.+)$/.exec(req.headers.get("authorization") ?? "")?.[1];
  return (token && DEMO_TOKENS[token]) || null;
}

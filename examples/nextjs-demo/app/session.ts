// Stand-in for your auth client's token getter (Clerk's getToken(), a
// refreshed OAuth token…). httpStore() calls it before every request.
export async function getAccessToken(): Promise<string> {
  return "demo-token";
}

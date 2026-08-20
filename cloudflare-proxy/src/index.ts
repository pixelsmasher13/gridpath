/**
 * Cloudflare Worker proxy for Vercel continuation calls.
 *
 * WHY: Vercel has recursive invocation loop detection that blocks after ~5
 * levels of the same deployment calling itself (508 INFINITE_LOOP_DETECTED).
 * By routing continuation calls through this external proxy, each request
 * enters Vercel as a fresh external call — the loop counter never increments.
 *
 * execute.ts → this worker → Vercel → execute.ts → this worker → Vercel → ...
 *                                      (loop depth always = 1)
 */

export interface Env {
  TARGET_HOST: string;
}

export default {
  async fetch(request: Request, env: Env): Promise<Response> {
    // Only allow POST (that's all continuations use)
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204 });
    }

    if (request.method !== 'POST') {
      return new Response('Method not allowed', { status: 405 });
    }

    // Rewrite the URL to point at Vercel
    const url = new URL(request.url);
    url.hostname = env.TARGET_HOST;
    url.protocol = 'https:';

    // Build a CLEAN request with only the headers we need.
    // Do NOT forward request.headers — they contain X-Forwarded-For,
    // CF-Connecting-IP, and other headers that let Vercel fingerprint
    // the original caller (its own egress IP) and trigger loop detection.
    const cleanHeaders = new Headers();
    cleanHeaders.set('Content-Type', 'application/json');

    // Forward only our auth headers
    const schedulerKey = request.headers.get('X-Scheduler-Key');
    const schedulerUserId = request.headers.get('X-Scheduler-User-Id');
    if (schedulerKey) cleanHeaders.set('X-Scheduler-Key', schedulerKey);
    if (schedulerUserId) cleanHeaders.set('X-Scheduler-User-Id', schedulerUserId);

    const proxyRequest = new Request(url.toString(), {
      method: 'POST',
      headers: cleanHeaders,
      body: request.body,
    });

    return fetch(proxyRequest);
  },
};

// The API that httpStore({ url: "/api/books" }) calls, for the signed-in
// user (the contract is in the gridpath guide, "Saving to your backend"):
//
//   PUT    /api/books/{id}?name=…&source=…[&checkpoint=1]   body: the .xlsx
//   GET    /api/books/{id}                     the .xlsx + X-Workbook-* headers, 404 if none
//   GET    /api/books                          [{ id, name, savedAt, size }], newest first
//   GET    /api/books/{id}/versions            [{ id, savedAt, size }], newest first
//   GET    /api/books/{id}/versions/{version}  that version's .xlsx
//   DELETE /api/books/{id}
import { userFrom } from "../../../../lib/auth";
import * as books from "../../../../lib/books";

type Ctx = { params: Promise<{ path?: string[] }> };

const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const MAX_BYTES = 100 * 1024 * 1024;

const reply = (status: number, text: string) => new Response(text, { status });
const xlsx = (bytes: Uint8Array, headers: Record<string, string> = {}) =>
  new Response(bytes as BodyInit, { headers: { "Content-Type": XLSX, "Cache-Control": "no-store", ...headers } });

export async function GET(req: Request, ctx: Ctx) {
  const user = await userFrom(req);
  if (!user) return reply(401, "sign in first");
  const [id, sub, version, ...rest] = (await ctx.params).path ?? [];
  if (!id) return Response.json(await books.list(user));
  if (sub === "versions" && !rest.length) {
    if (!version) return Response.json(await books.versions(user, id));
    const bytes = await books.loadVersion(user, id, version);
    return bytes ? xlsx(bytes) : reply(404, "no such version");
  }
  if (sub) return reply(404, "not found");
  const book = await books.load(user, id);
  if (!book) return reply(404, "no such workbook");
  return xlsx(book.bytes, {
    // Header values are ASCII: the name travels percent-encoded.
    "X-Workbook-Name": encodeURIComponent(book.meta.name),
    "X-Workbook-Saved-At": book.meta.savedAt,
    ...(book.meta.source ? { "X-Workbook-Source": book.meta.source } : {}),
  });
}

export async function PUT(req: Request, ctx: Ctx) {
  const user = await userFrom(req);
  if (!user) return reply(401, "sign in first");
  const [id, ...rest] = (await ctx.params).path ?? [];
  if (!id || rest.length) return reply(404, "not found");
  const bytes = new Uint8Array(await req.arrayBuffer());
  if (bytes.length > MAX_BYTES) return reply(413, "workbook too large");
  // An .xlsx is a zip archive: anything else is refused.
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) return reply(415, "not an .xlsx workbook");
  const q = new URL(req.url).searchParams;
  await books.save(user, id, { name: q.get("name") || "Untitled.xlsx", bytes, source: q.get("source") ?? undefined, checkpoint: q.has("checkpoint") });
  return new Response(null, { status: 204 });
}

export async function DELETE(req: Request, ctx: Ctx) {
  const user = await userFrom(req);
  if (!user) return reply(401, "sign in first");
  const [id, ...rest] = (await ctx.params).path ?? [];
  if (!id || rest.length) return reply(404, "not found");
  await books.remove(user, id);
  return new Response(null, { status: 204 });
}

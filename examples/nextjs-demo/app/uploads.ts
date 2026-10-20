// A file the user picked, handed to its workbook page across a client-side
// navigation (module state survives it). Read when the page mounts,
// forgotten once the spreadsheet has opened it.
const picked = new Map<string, { name: string; bytes: Uint8Array }>();

export const keepUpload = (id: string, file: { name: string; bytes: Uint8Array }) => void picked.set(id, file);
export const uploadFor = (id: string) => picked.get(id) ?? null;
export const forgetUpload = (id: string) => void picked.delete(id);

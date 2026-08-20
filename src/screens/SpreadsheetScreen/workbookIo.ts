import { invoke } from "@tauri-apps/api/core";

export async function readWorkbookBytes(path: string): Promise<Uint8Array> {
  const b64 = await invoke<string>("read_workbook_file", { path });
  return base64ToBytes(b64);
}

export async function writeWorkbookBytes(path: string, bytes: Uint8Array): Promise<void> {
  await invoke("write_workbook_file", {
    path,
    bytesB64: bytesToBase64(bytes),
  });
}

/**
 * Surgical save: apply a cell/format patch to the workbook on disk without
 * re-serializing the package (drawings, external links, pivots etc. survive
 * byte-identical). Throws with a message starting "BASE_CHANGED:" when the
 * file was modified outside GridPath, or "PATCH_FAILED:" when the patcher
 * couldn't apply — callers decide whether to fall back to the full export.
 */
export async function saveWorkbookPatched(path: string, patchJson: string): Promise<void> {
  await invoke("save_workbook_patched", { path, patchJson });
}

/**
 * Save As, surgically: clone the source package and apply the patch to the
 * copy — the new file keeps every part of the original byte-identical.
 */
export async function saveWorkbookPatchedAs(
  sourcePath: string,
  targetPath: string,
  patchJson: string,
): Promise<void> {
  await invoke("save_workbook_patched_as", { sourcePath, targetPath, patchJson });
}

/**
 * Persist a Univer-native snapshot (the lossless JSON shape) for an
 * untitled draft. Routes through the same byte-write Tauri command we use
 * for xlsx — we just UTF-8 encode the JSON and let the existing path
 * resolution land it under app_data_dir/untitled_sessions/. The `.gpsnap`
 * suffix differentiates these from real .xlsx files so a future cleanup
 * sweep can tell them apart.
 */
export async function writeUntitledSnapshot(workbookPath: string, snapshot: unknown): Promise<void> {
  const json = JSON.stringify(snapshot);
  const bytes = new TextEncoder().encode(json);
  await invoke("write_workbook_file", {
    path: `${workbookPath}.gpsnap`,
    bytesB64: bytesToBase64(bytes),
  });
}

/**
 * Inverse of writeUntitledSnapshot — returns null if the snapshot file
 * doesn't exist (e.g. fresh untitled tab on a clean install, or a draft
 * predating the JSON-snapshot change which only has the xlsx file).
 */
export async function readUntitledSnapshot(workbookPath: string): Promise<unknown | null> {
  try {
    const b64 = await invoke<string>("read_workbook_file", { path: `${workbookPath}.gpsnap` });
    const bytes = base64ToBytes(b64);
    const json = new TextDecoder().decode(bytes);
    return JSON.parse(json);
  } catch {
    return null;
  }
}

function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function bytesToBase64(bytes: Uint8Array): string {
  let bin = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    bin += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + chunk)) as number[]);
  }
  return btoa(bin);
}

import { readFileSync, statSync, openSync, readSync, closeSync } from "fs";
import { basename } from "path";
import mime from "mime-types";
import { graphRequest } from "./graph.js";

export const INLINE_LIMIT_BYTES = 3 * 1024 * 1024;
const UPLOAD_CHUNK_BYTES = 5 * 1024 * 1024;

export interface LoadedAttachment {
  path: string;
  name: string;
  contentType: string;
  size: number;
}

export function loadAttachment(path: string): LoadedAttachment {
  const stat = statSync(path);
  if (!stat.isFile()) throw new Error(`Not a file: ${path}`);
  const name = basename(path);
  const contentType = mime.lookup(name) || "application/octet-stream";
  return { path, name, contentType, size: stat.size };
}

export function toFileAttachmentJson(att: LoadedAttachment): object {
  const bytes = readFileSync(att.path);
  return {
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: att.name,
    contentType: att.contentType,
    contentBytes: bytes.toString("base64"),
  };
}

export async function uploadLargeAttachment(messageId: string, att: LoadedAttachment): Promise<void> {
  const session = await graphRequest(`/me/messages/${messageId}/attachments/createUploadSession`, {
    method: "POST",
    body: JSON.stringify({
      AttachmentItem: {
        attachmentType: "file",
        name: att.name,
        size: att.size,
        contentType: att.contentType,
      },
    }),
  });

  const uploadUrl: string = session.uploadUrl;
  if (!uploadUrl) throw new Error("Upload session missing uploadUrl");

  const fd = openSync(att.path, "r");
  try {
    let offset = 0;
    while (offset < att.size) {
      const end = Math.min(offset + UPLOAD_CHUNK_BYTES, att.size) - 1;
      const chunkLen = end - offset + 1;
      const buf = Buffer.alloc(chunkLen);
      readSync(fd, buf, 0, chunkLen, offset);

      const res = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": String(chunkLen),
          "Content-Range": `bytes ${offset}-${end}/${att.size}`,
          "Content-Type": "application/octet-stream",
        },
        body: buf,
      });

      if (!res.ok && res.status !== 201 && res.status !== 200) {
        const errBody = await res.text();
        throw new Error(`Upload chunk failed ${res.status}: ${errBody}`);
      }

      offset = end + 1;
    }
  } finally {
    closeSync(fd);
  }
}

export interface AttachedSummary {
  name: string;
  size: number;
  id?: string;
}

export async function attachAllToDraft(messageId: string, paths: string[]): Promise<AttachedSummary[]> {
  const summary: AttachedSummary[] = [];
  for (const p of paths) {
    const att = loadAttachment(p);
    if (att.size <= INLINE_LIMIT_BYTES) {
      const res = await graphRequest(`/me/messages/${messageId}/attachments`, {
        method: "POST",
        body: JSON.stringify(toFileAttachmentJson(att)),
      });
      summary.push({ name: att.name, size: att.size, id: res?.id });
    } else {
      await uploadLargeAttachment(messageId, att);
      summary.push({ name: att.name, size: att.size });
    }
  }
  return summary;
}

export function hasLarge(paths: string[]): boolean {
  return paths.some((p) => statSync(p).size > INLINE_LIMIT_BYTES);
}

export function loadAll(paths: string[]): LoadedAttachment[] {
  return paths.map(loadAttachment);
}

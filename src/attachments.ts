// Image / PDF attachments (invoices, receipts, screenshots) sent to Gemini as inlineData.

export interface Attachment {
  name: string;
  mimeType: string;
  data: string; // base64 without prefix
  previewUrl?: string; // small data URL for thumbnails (images only)
  bytes: number;
}

const MAX_IMAGE_SIDE = 2400; // keeps small invoice text readable while limiting payload
export const MAX_TOTAL_BYTES = 18 * 1024 * 1024; // inline request limit is ~20MB

function readAsDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as string);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });
}

function splitDataUrl(url: string): { mimeType: string; data: string } {
  const m = url.match(/^data:([^;]+);base64,(.*)$/);
  if (!m) throw new Error("无法读取文件");
  return { mimeType: m[1], data: m[2] };
}

async function makeThumb(bitmap: ImageBitmap): Promise<string> {
  const s = Math.min(1, 96 / Math.max(bitmap.width, bitmap.height));
  const c = document.createElement("canvas");
  c.width = Math.round(bitmap.width * s);
  c.height = Math.round(bitmap.height * s);
  c.getContext("2d")!.drawImage(bitmap, 0, 0, c.width, c.height);
  return c.toDataURL("image/jpeg", 0.7);
}

export async function fileToAttachment(file: File): Promise<Attachment> {
  const name = file.name || "粘贴的图片.png";
  if (file.type === "application/pdf" || /\.pdf$/i.test(name)) {
    const { data } = splitDataUrl(await readAsDataUrl(file));
    return { name, mimeType: "application/pdf", data, bytes: file.size };
  }
  if (file.type.startsWith("image/") || /\.(png|jpe?g|webp|heic|heif|bmp|gif)$/i.test(name)) {
    try {
      const bitmap = await createImageBitmap(file);
      const scale = Math.min(1, MAX_IMAGE_SIDE / Math.max(bitmap.width, bitmap.height));
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(bitmap.width * scale);
      canvas.height = Math.round(bitmap.height * scale);
      const g = canvas.getContext("2d")!;
      g.fillStyle = "#fff"; // flatten transparency
      g.fillRect(0, 0, canvas.width, canvas.height);
      g.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
      const { mimeType, data } = splitDataUrl(canvas.toDataURL("image/jpeg", 0.9));
      const previewUrl = await makeThumb(bitmap);
      bitmap.close();
      return { name, mimeType, data, previewUrl, bytes: Math.round((data.length * 3) / 4) };
    } catch {
      // Browser can't decode (e.g. HEIC): send the original file, Gemini can read it.
      const { data } = splitDataUrl(await readAsDataUrl(file));
      const mimeType = file.type || (/\.heif$/i.test(name) ? "image/heif" : "image/heic");
      return { name, mimeType, data, bytes: file.size };
    }
  }
  throw new Error(`不支持的文件类型：${name}（支持图片和 PDF）`);
}

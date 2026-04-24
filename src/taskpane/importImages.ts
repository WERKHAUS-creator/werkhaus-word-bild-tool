import type { ImageItem, StoredMeta } from "./types";

// Stabiler Kernbereich:
// Datei- und Ordnerimport muessen immer in dieselbe interne Bildstruktur muenden.
// Bitte nur bei nachgewiesenem Fehler aendern.

export interface ImageImportPersistence {
  getMeta(hash: string): StoredMeta;
  setMeta(hash: string, meta: StoredMeta): void;
}

export interface ImportImageFilesResult {
  imageFileCount: number;
  invalidFileCount: number;
  invalidFileNames: string[];
  addedCount: number;
  skippedCount: number;
  failedCount: number;
  failedFileNames: string[];
  items: ImageItem[];
}

const SUPPORTED_IMAGE_MIME_TYPES = new Set([
  "image/jpeg",
  "image/png",
  "image/webp",
  "image/tiff",
  "image/gif",
  "image/bmp",
]);

const SUPPORTED_IMAGE_EXTENSIONS = new Set([
  ".jpg",
  ".jpeg",
  ".png",
  ".webp",
  ".tif",
  ".tiff",
  ".gif",
  ".bmp",
]);

export const IMAGE_FILE_INPUT_ACCEPT = Array.from(SUPPORTED_IMAGE_EXTENSIONS).join(",");

export function isSupportedImageFile(file: File): boolean {
  const normalizedType = (file.type || "").toLowerCase();
  if (SUPPORTED_IMAGE_MIME_TYPES.has(normalizedType)) {
    return true;
  }

  const normalizedName = file.name.toLowerCase();
  return Array.from(SUPPORTED_IMAGE_EXTENSIONS).some((extension) =>
    normalizedName.endsWith(extension)
  );
}

export async function importImageFiles(
  files: FileList | File[],
  existingItems: ImageItem[],
  persistence: ImageImportPersistence
): Promise<ImportImageFilesResult> {
  // Stabiler Kernbereich: derselbe Importpfad fuer Auswahl, Ordner und Drop.
  const allFiles = Array.from(files);
  const imageFiles = allFiles.filter((file) => isSupportedImageFile(file));
  const nextItems = [...existingItems];
  const invalidFileCount = allFiles.length - imageFiles.length;
  const invalidFileNames = allFiles
    .filter((file) => !isSupportedImageFile(file))
    .map((file) => file.name);

  if (imageFiles.length === 0) {
    return {
      imageFileCount: 0,
      invalidFileCount,
      invalidFileNames,
      addedCount: 0,
      skippedCount: 0,
      failedCount: 0,
      failedFileNames: [],
      items: nextItems,
    };
  }

  let addedCount = 0;
  let skippedCount = 0;
  let failedCount = 0;
  const failedFileNames: string[] = [];

  for (const file of imageFiles) {
    try {
      let hashHex = "";
      let buffer: ArrayBuffer | null = null;

      try {
        buffer = await file.arrayBuffer();
        const hashBuffer = await crypto.subtle.digest("SHA-256", buffer);
        const hashArray = Array.from(new Uint8Array(hashBuffer));
        hashHex = hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
      } catch (err) {
        console.warn("Fehler beim Berechnen des Datei-Hashes:", err);
        hashHex = makeFileKey(file);
      }

      const exists = nextItems.some((item) => item.hash === hashHex);
      if (exists) {
        skippedCount += 1;
        continue;
      }

      const dataUrl = await readFileAsDataUrl(file);
      const key = makeFileKey(file);
      const stored = persistence.getMeta(hashHex);
      const exifParsed = buffer ? parseExif(buffer) : {};
      const imageDimensions = await getImageDimensions(dataUrl);

      const pos =
        stored.position && typeof stored.position === "number"
          ? stored.position
          : nextItems.length + 1;

      nextItems.push({
        id: createId(),
        key,
        name: file.name,
        size: file.size,
        base64: dataUrlToBase64(dataUrl),
        previewUrl: dataUrl,
        hash: hashHex,
        caption: stored.caption || "",
        position: pos,
        selected: stored.selected ?? true,
        includeCaptionInWord: stored.includeCaptionInWord ?? true,
        exif: {
          dateTimeOriginal: exifParsed.dateTimeOriginal,
          model: exifParsed.model,
          width: imageDimensions?.width,
          height: imageDimensions?.height,
        },
      });

      if (!stored.importedAt) {
        persistence.setMeta(hashHex, {
          ...(stored || {}),
          importedAt: new Date().toISOString(),
          position: pos,
          selected: stored.selected ?? true,
          includeCaptionInWord: stored.includeCaptionInWord ?? true,
        });
      }

      addedCount += 1;
    } catch (error) {
      failedCount += 1;
      failedFileNames.push(file.name);
      console.warn(`Fehler beim Import von ${file.name}:`, error);
    }
  }

  return {
    imageFileCount: imageFiles.length,
    invalidFileCount,
    invalidFileNames,
    addedCount,
    skippedCount,
    failedCount,
    failedFileNames,
    items: nextItems,
  };
}

function createId(): string {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function makeFileKey(file: File): string {
  return `${file.name}__${file.size}__${file.lastModified}`;
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      if (typeof reader.result === "string") {
        resolve(reader.result);
      } else {
        reject(new Error("Datei konnte nicht gelesen werden."));
      }
    };

    reader.onerror = () => reject(new Error("Fehler beim Lesen der Datei."));
    reader.readAsDataURL(file);
  });
}

function dataUrlToBase64(dataUrl: string): string {
  const commaIndex = dataUrl.indexOf(",");
  if (commaIndex === -1) {
    throw new Error("Ungültiges Data-URL-Format.");
  }

  return dataUrl.substring(commaIndex + 1);
}

async function getImageDimensions(
  dataUrl: string
): Promise<{ width: number; height: number } | undefined> {
  if (typeof Image === "undefined") {
    return undefined;
  }

  return new Promise((resolve) => {
    const image = new Image();

    image.onload = () => {
      const width = image.naturalWidth || image.width || 0;
      const height = image.naturalHeight || image.height || 0;

      if (width > 0 && height > 0) {
        resolve({ width, height });
      } else {
        resolve(undefined);
      }
    };

    image.onerror = () => resolve(undefined);
    image.src = dataUrl;
  });
}

function parseExif(arrayBuffer: ArrayBuffer): { dateTimeOriginal?: string; model?: string } {
  try {
    const view = new DataView(arrayBuffer);
    if (view.getUint16(0) !== 0xffd8) return {};

    let offset = 2;

    while (offset < view.byteLength) {
      if (view.getUint8(offset) !== 0xff) break;

      const marker = view.getUint8(offset + 1);
      const size = view.getUint16(offset + 2);

      if (marker === 0xe1) {
        const exifHeader = offset + 4;
        const exifStr = String.fromCharCode(
          view.getUint8(exifHeader),
          view.getUint8(exifHeader + 1),
          view.getUint8(exifHeader + 2),
          view.getUint8(exifHeader + 3),
          view.getUint8(exifHeader + 4),
          view.getUint8(exifHeader + 5)
        );

        if (exifStr !== "Exif\0\0") return {};

        const tiffOffset = exifHeader + 6;
        const endianMark = view.getUint16(tiffOffset);
        const little = endianMark === 0x4949;

        const getUint16 = (off: number) => view.getUint16(off, little);
        const getUint32 = (off: number) => view.getUint32(off, little);

        const firstIFDOffset = getUint32(tiffOffset + 4) + tiffOffset;
        const entries = getUint16(firstIFDOffset);

        let dateTime: string | undefined;
        let model: string | undefined;

        for (let i = 0; i < entries; i++) {
          const entryOffset = firstIFDOffset + 2 + i * 12;
          const tag = getUint16(entryOffset);
          const count = getUint32(entryOffset + 4);
          const valueOffset = getUint32(entryOffset + 8);

          if (tag === 0x0110) {
            const strOffset = count > 4 ? tiffOffset + valueOffset : entryOffset + 8;
            let s = "";
            for (let k = 0; k < count - 1; k++) {
              const c = view.getUint8(strOffset + k);
              if (c === 0) break;
              s += String.fromCharCode(c);
            }
            model = s;
          }

          if (tag === 0x8769) {
            const exifIFDOffset = tiffOffset + valueOffset;
            const exifEntries = getUint16(exifIFDOffset);

            for (let j = 0; j < exifEntries; j++) {
              const eOff = exifIFDOffset + 2 + j * 12;
              const etag = getUint16(eOff);
              const ecount = getUint32(eOff + 4);
              const evalOff = getUint32(eOff + 8);

              if (etag === 0x9003) {
                const strOff = ecount > 4 ? tiffOffset + evalOff : eOff + 8;
                let ds = "";
                for (let k = 0; k < ecount - 1; k++) {
                  const c = view.getUint8(strOff + k);
                  if (c === 0) break;
                  ds += String.fromCharCode(c);
                }
                dateTime = ds;
              }
            }
          }
        }

        return { dateTimeOriginal: dateTime, model };
      }

      offset += 2 + size;
    }
  } catch (e) {
    console.warn("EXIF parse error", e);
  }

  return {};
}

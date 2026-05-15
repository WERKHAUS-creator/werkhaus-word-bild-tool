import type { ImageItem, StoredMeta } from "./types";

// Stabiler Kernbereich:
// Datei- und Ordnerimport muessen immer in dieselbe interne Bildstruktur muenden.
// Bitte nur bei nachgewiesenem Fehler aendern.

export interface ImageImportPersistence {
  getMeta(hash: string): StoredMeta;
  getMetaKeys?: () => string[];
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
  const allFiles = Array.from(files);
  const invalidFiles = allFiles.filter((file) => !isSupportedImageFile(file));
  const imageFiles = allFiles.filter((file) => !invalidFiles.includes(file));
  const nextItems = [...existingItems];
  const knownKeys = new Set(nextItems.map((item) => item.key || item.hash));
  const metaKeys = persistence.getMetaKeys ? persistence.getMetaKeys() : [];
  const hasLegacyHashEntries = metaKeys.some(isLegacySha256Key);
  const invalidFileCount = allFiles.length - imageFiles.length;
  const invalidFileNames = invalidFiles.map((file) => file.name);

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
      const key = makeFileKey(file);
      const hashHex = key;

      if (knownKeys.has(key)) {
        skippedCount += 1;
        continue;
      }

      const dataUrl = await readFileAsDataUrl(file);
      const dimensions = await getImageDimensions(dataUrl);
      const stored = persistence.getMeta(hashHex);
      const migratedLegacyCaption =
        hasLegacyHashEntries && !hasCaption(stored)
          ? await getLegacyCaptionMeta(file, persistence)
          : undefined;
      const effectiveStored = migratedLegacyCaption
        ? { ...stored, ...migratedLegacyCaption }
        : stored;

      if (migratedLegacyCaption) {
        // Feldweise Migration: nur fehlende Caption uebernehmen, keine weiteren Metadaten ueberschreiben.
        persistence.setMeta(hashHex, effectiveStored);
      }

      const pos =
        effectiveStored.position && typeof effectiveStored.position === "number"
          ? effectiveStored.position
          : nextItems.length + 1;

      nextItems.push({
        id: createId(),
        key,
        name: file.name,
        relativePath: getFileRelativePath(file),
        lastModified: file.lastModified,
        size: file.size,
        base64: dataUrlToBase64(dataUrl),
        previewUrl: dataUrl,
        hash: hashHex,
        caption: effectiveStored.caption || "",
        position: pos,
        selected: effectiveStored.selected ?? true,
        includeCaptionInWord: effectiveStored.includeCaptionInWord ?? true,
        exif: dimensions
          ? {
              width: dimensions.width,
              height: dimensions.height,
            }
          : undefined,
      });
      knownKeys.add(key);

      if (!stored.importedAt) {
        persistence.setMeta(hashHex, {
          ...(effectiveStored || {}),
          importedAt: new Date().toISOString(),
          position: pos,
          selected: effectiveStored.selected ?? true,
          includeCaptionInWord: effectiveStored.includeCaptionInWord ?? true,
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

function getFileRelativePath(file: File): string | undefined {
  const fileWithRelativePath = file as File & {
    relativePath?: string;
    webkitRelativePath?: string;
  };
  const relativePath =
    fileWithRelativePath.relativePath || fileWithRelativePath.webkitRelativePath || "";

  if (relativePath.trim().length > 0) {
    return relativePath;
  }

  return file.name || undefined;
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

function hasCaption(meta: StoredMeta | undefined): boolean {
  return Boolean(meta && typeof meta.caption === "string" && meta.caption.trim().length > 0);
}

async function getLegacyCaptionMeta(
  file: File,
  persistence: ImageImportPersistence
): Promise<Pick<StoredMeta, "caption" | "comment"> | undefined> {
  try {
    const legacyHash = await computeLegacySha256(file);
    if (!legacyHash) {
      return undefined;
    }

    const legacyStored = persistence.getMeta(legacyHash);
    if (!hasCaption(legacyStored)) {
      return undefined;
    }

    return {
      caption: legacyStored.caption,
      comment: legacyStored.comment,
    };
  } catch (error) {
    console.warn("Konnte alte Bild-Captions nicht migrieren:", error);
    return undefined;
  }
}

function isLegacySha256Key(key: string): boolean {
  return /^[a-f0-9]{64}$/i.test(key);
}

async function computeLegacySha256(file: File): Promise<string | undefined> {
  if (typeof crypto === "undefined" || !crypto.subtle) {
    return undefined;
  }

  const buffer = await file.arrayBuffer();
  const hashBuffer = await crypto.subtle.digest("SHA-256", buffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((value) => value.toString(16).padStart(2, "0")).join("");
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

import type { ImageItem, SortMode } from "./types";

export const PROJECT_FILE_NAME = "bilddaten.json";
const PROJECT_TYPE = "WERKHAUS-Bilddaten";
const PROJECT_SCHEMA_VERSION = 1;

export interface BilddatenProjectImage {
  relativePath?: string;
  filename: string;
  position: number;
  imageNumber: string;
  caption?: string;
  active: boolean;
  includeCaptionInWord: boolean;
  fileSize?: number;
  lastModified?: number;
  exifDateTaken?: string;
}

export interface BilddatenProjectFile {
  schemaVersion: number;
  projectType: string;
  createdAt: string;
  updatedAt: string;
  sortMode: SortMode;
  images: BilddatenProjectImage[];
}

export interface ProjectMatch {
  image: BilddatenProjectImage;
  itemIndex: number;
}

export interface ProjectMatchResult {
  matches: ProjectMatch[];
  unmatchedProjectImages: BilddatenProjectImage[];
  unmatchedItemIndexes: number[];
}

export function buildBilddatenProjectFile(
  items: ImageItem[],
  sortMode: SortMode,
  createdAt?: string
): BilddatenProjectFile {
  const now = new Date().toISOString();
  return {
    schemaVersion: PROJECT_SCHEMA_VERSION,
    projectType: PROJECT_TYPE,
    createdAt: createdAt || now,
    updatedAt: now,
    sortMode,
    images: items.map((item, index) => buildProjectImage(item, index + 1)),
  };
}

export function serializeBilddatenProjectFile(projectFile: BilddatenProjectFile): string {
  return JSON.stringify(projectFile, null, 2);
}

export function parseBilddatenProjectFile(rawText: string): BilddatenProjectFile | undefined {
  try {
    const parsed = JSON.parse(rawText) as Partial<BilddatenProjectFile>;

    if (!parsed || parsed.schemaVersion !== PROJECT_SCHEMA_VERSION) {
      return undefined;
    }

    if (parsed.projectType !== PROJECT_TYPE) {
      return undefined;
    }

    if (!Array.isArray(parsed.images)) {
      return undefined;
    }

    const sortMode: SortMode =
      parsed.sortMode === "exifDate" || parsed.sortMode === "name" || parsed.sortMode === "custom"
        ? parsed.sortMode
        : "custom";

    const images = parsed.images
      .map((image) => normalizeProjectImage(image))
      .filter((image): image is BilddatenProjectImage => Boolean(image));

    return {
      schemaVersion: PROJECT_SCHEMA_VERSION,
      projectType: PROJECT_TYPE,
      createdAt: normalizeString(parsed.createdAt) || new Date().toISOString(),
      updatedAt: normalizeString(parsed.updatedAt) || new Date().toISOString(),
      sortMode,
      images,
    };
  } catch (error) {
    console.warn("Konnte bilddaten.json nicht parsen:", error);
    return undefined;
  }
}

export function matchProjectImagesToItems(
  projectFile: BilddatenProjectFile,
  items: ImageItem[]
): ProjectMatchResult {
  const usedItemIndexes = new Set<number>();
  const matches: ProjectMatch[] = [];
  const unmatchedProjectImages: BilddatenProjectImage[] = [];

  for (const projectImage of projectFile.images) {
    const itemIndex = findBestMatchingItemIndex(projectImage, items, usedItemIndexes);

    if (itemIndex === -1) {
      unmatchedProjectImages.push(projectImage);
      continue;
    }

    usedItemIndexes.add(itemIndex);
    matches.push({
      image: projectImage,
      itemIndex,
    });
  }

  const unmatchedItemIndexes: number[] = [];
  for (let index = 0; index < items.length; index += 1) {
    if (!usedItemIndexes.has(index)) {
      unmatchedItemIndexes.push(index);
    }
  }

  return {
    matches,
    unmatchedProjectImages,
    unmatchedItemIndexes,
  };
}

function buildProjectImage(item: ImageItem, position: number): BilddatenProjectImage {
  const fileNumber = String(position).padStart(3, "0");

  return {
    relativePath: normalizeOptionalText(item.relativePath),
    filename: item.name,
    position,
    imageNumber: fileNumber,
    caption: normalizeOptionalText(item.caption),
    active: item.selected !== false,
    includeCaptionInWord: item.includeCaptionInWord !== false,
    fileSize: Number.isFinite(item.size) ? item.size : undefined,
    lastModified: Number.isFinite(item.lastModified ?? NaN) ? item.lastModified : undefined,
    exifDateTaken: normalizeOptionalText(item.exif?.dateTimeOriginal),
  };
}

function normalizeProjectImage(
  image: Partial<BilddatenProjectImage>
): BilddatenProjectImage | undefined {
  const filename = normalizeString(image.filename);
  const position = Number(image.position);
  const imageNumber =
    normalizeString(image.imageNumber) ||
    String(Math.max(1, Math.floor(position || 1))).padStart(3, "0");

  if (!filename) {
    return undefined;
  }

  return {
    relativePath: normalizeOptionalText(image.relativePath),
    filename,
    position: Number.isFinite(position) && position > 0 ? Math.floor(position) : 0,
    imageNumber,
    caption: normalizeOptionalText(image.caption),
    active: image.active !== false,
    includeCaptionInWord: image.includeCaptionInWord !== false,
    fileSize: normalizeNumber(image.fileSize),
    lastModified: normalizeNumber(image.lastModified),
    exifDateTaken: normalizeOptionalText(image.exifDateTaken),
  };
}

function findBestMatchingItemIndex(
  projectImage: BilddatenProjectImage,
  items: ImageItem[],
  usedItemIndexes: Set<number>
): number {
  const matchPathAndMeta = (item: ImageItem) =>
    hasMatchingRelativePath(projectImage, item) && hasMatchingSizeMeta(projectImage, item);
  const matchPath = (item: ImageItem) => hasMatchingRelativePath(projectImage, item);
  const matchNameAndSizeModified = (item: ImageItem) =>
    hasMatchingFilename(projectImage, item) && hasMatchingSizeAndModified(projectImage, item);
  const matchNameAndSize = (item: ImageItem) =>
    hasMatchingFilename(projectImage, item) && hasMatchingFileSize(projectImage, item);
  const matchName = (item: ImageItem) => hasMatchingFilename(projectImage, item);

  const matchLevels: Array<(item: ImageItem) => boolean> = [
    matchPathAndMeta,
    matchPath,
    matchNameAndSizeModified,
    matchNameAndSize,
    matchName,
  ];

  for (const matchesLevel of matchLevels) {
    for (let index = 0; index < items.length; index += 1) {
      if (usedItemIndexes.has(index)) {
        continue;
      }

      if (matchesLevel(items[index])) {
        return index;
      }
    }
  }

  return -1;
}

function hasMatchingRelativePath(projectImage: BilddatenProjectImage, item: ImageItem): boolean {
  const projectRelativePath = normalizePath(projectImage.relativePath);
  const itemRelativePath = normalizePath(item.relativePath);

  return Boolean(
    projectRelativePath && itemRelativePath && projectRelativePath === itemRelativePath
  );
}

function hasMatchingFilename(projectImage: BilddatenProjectImage, item: ImageItem): boolean {
  return normalizeName(projectImage.filename) === normalizeName(item.name);
}

function hasMatchingFileSize(projectImage: BilddatenProjectImage, item: ImageItem): boolean {
  return normalizeNumber(projectImage.fileSize) === normalizeNumber(item.size);
}

function hasMatchingSizeAndModified(projectImage: BilddatenProjectImage, item: ImageItem): boolean {
  return (
    hasMatchingFileSize(projectImage, item) &&
    normalizeNumber(projectImage.lastModified) === normalizeNumber(item.lastModified)
  );
}

function hasMatchingSizeMeta(projectImage: BilddatenProjectImage, item: ImageItem): boolean {
  return (
    hasMatchingRelativePath(projectImage, item) &&
    hasMatchingFileSize(projectImage, item) &&
    normalizeNumber(projectImage.lastModified) === normalizeNumber(item.lastModified)
  );
}

function normalizePath(value?: string): string {
  return normalizeString(value).replace(/\\/g, "/").toLowerCase();
}

function normalizeName(value?: string): string {
  return normalizeString(value).toLowerCase();
}

function normalizeString(value?: string): string {
  return typeof value === "string" ? value.trim() : "";
}

function normalizeOptionalText(value: unknown): string | undefined {
  const normalized = typeof value === "string" ? value.trim() : "";
  return normalized.length > 0 ? normalized : undefined;
}

function normalizeNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

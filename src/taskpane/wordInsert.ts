import { USABLE_PAGE_HEIGHT_CM, USABLE_PAGE_WIDTH_CM, WORD_PT_PER_INCH } from "./constants";
import { getResolvedCaptionText, shouldInsertCaption } from "./captions";
import type { ImageItem } from "./types";

const CM_TO_POINTS = WORD_PT_PER_INCH / 2.54;

interface InsertSingleImageOptions {
  includeCaption?: boolean;
}

// Stabiler Kernbereich:
// Einzelfoto-Einfuegen und Caption-Verhalten bleiben unveraendert.
// Nicht mit Kontaktbogen- oder Importexperimenten vermischen.

export async function insertSingleImageAtSelection(
  item: ImageItem,
  insertSizeCm: number,
  options: InsertSingleImageOptions = {}
) {
  const base64 = normalizeBase64ForWord(item.base64);
  const size = getScaledImageSize(item, insertSizeCm);
  const figureNumber = resolveSidebarFigureNumber(item);

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const shouldIncludeCaption =
      typeof options.includeCaption === "boolean"
        ? options.includeCaption
        : shouldInsertCaption(item);

    const pic = selection.insertInlinePictureFromBase64(base64, Word.InsertLocation.replace);

    try {
      pic.width = size.width;
      pic.height = size.height;
    } catch (e) {
      console.warn("Konnte Bildgröße nicht setzen:", e);
    }

    await context.sync();

    let continuationRange: Word.Range;
    if (shouldIncludeCaption) {
      const captionText = getResolvedCaptionText(item, figureNumber);
      continuationRange = insertCaptionAfterPicture(pic, captionText);
    } else {
      continuationRange = pic.getRange(Word.RangeLocation.end);
    }

    await context.sync();

    await placeCursorAfterRange(context, continuationRange);
  });
}

function resolveSidebarFigureNumber(item: ImageItem): number {
  // Die Bildnummer soll exakt der Reihenfolge in der Seitenleiste folgen.
  const position = Number(item.position);

  if (Number.isFinite(position) && position > 0) {
    return Math.floor(position);
  }

  return 1;
}

function getScaledImageSize(item: ImageItem, targetSizeCm: number) {
  let origW = item.exif?.width || 0;
  let origH = item.exif?.height || 0;

  if (!origW || !origH) {
    origW = 800;
    origH = 600;
  }

  // Der Reglerwert beschreibt die laengste Seite des Bildes.
  const finalMaxCm = Math.max(
    0.1,
    Math.min(targetSizeCm, USABLE_PAGE_WIDTH_CM, USABLE_PAGE_HEIGHT_CM)
  );
  const maxSidePoints = finalMaxCm * CM_TO_POINTS;

  let targetWPoints = maxSidePoints;
  let targetHPoints = maxSidePoints;

  if (origW > 0 && origH > 0) {
    if (origW >= origH) {
      targetWPoints = maxSidePoints;
      targetHPoints = Math.max(1, Math.round(targetWPoints * (origH / origW)));
    } else {
      targetHPoints = maxSidePoints;
      targetWPoints = Math.max(1, Math.round(targetHPoints * (origW / origH)));
    }
  }

  return { width: targetWPoints, height: targetHPoints };
}

function insertCaptionAfterPicture(picture: Word.InlinePicture, captionText: string): Word.Range {
  const pictureEndRange = picture.getRange(Word.RangeLocation.end);
  const captionParagraph = pictureEndRange.insertParagraph(captionText, Word.InsertLocation.after);

  // Word bietet hier ohne zusaetzlichen Container keine echte feste Caption-Blockbreite.
  // Wir begrenzen deshalb den Absatz ueber den rechten Einzug auf die aktuelle Bildbreite.
  const remainingWidth = Math.max(0, USABLE_PAGE_WIDTH_CM * CM_TO_POINTS - picture.width);
  captionParagraph.leftIndent = 0;
  captionParagraph.rightIndent = remainingWidth;
  captionParagraph.spaceBefore = 0;
  captionParagraph.spaceAfter = 0;

  const spacerParagraph = captionParagraph.insertParagraph("", Word.InsertLocation.after);

  return spacerParagraph.getRange(Word.RangeLocation.after);
}

async function placeCursorAfterRange(
  context: Word.RequestContext,
  range: Word.Range
): Promise<void> {
  range.select("End");
  await context.sync();
}

function normalizeBase64ForWord(imageBase64OrDataUrl: string): string {
  const normalized = (imageBase64OrDataUrl || "").trim();

  if (!normalized) {
    throw new Error("Bilddaten fehlen oder sind leer.");
  }

  const payload = normalized.includes(",") ? normalized.split(",").pop() || "" : normalized;
  const cleanedPayload = payload.replace(/\s+/g, "").trim();

  if (!cleanedPayload) {
    throw new Error("Bilddaten konnten nicht als Base64 für Word aufbereitet werden.");
  }

  return cleanedPayload;
}

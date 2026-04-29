import type { CaptionMode, ImageItem } from "./types";

// Stabiler Kernbereich:
// Caption-Regeln und Caption-Entscheidung bleiben getrennt von Import und Word-Ausgabe.
// Nur bei fachlicher Regelanderung anpassen.

export function getResolvedCaptionText(item: ImageItem, figureNumber: number): string {
  const trimmedCaption = item.caption?.trim() || "";

  if (trimmedCaption.length > 0) {
    return `Abbildung ${figureNumber}: ${trimmedCaption}`;
  }

  return `Abbildung ${figureNumber}:`;
}

export function shouldInsertCaption(item: ImageItem): boolean {
  return item.includeCaptionInWord !== false;
}

export function getFigureNumberText(figureNumber: number): string {
  return `Abbildung ${figureNumber}`;
}

export function getCaptionTextByMode(
  item: ImageItem,
  figureNumber: number,
  mode: CaptionMode
): string {
  if (mode === "numberOnly") {
    return getFigureNumberText(figureNumber);
  }

  return getResolvedCaptionText(item, figureNumber);
}

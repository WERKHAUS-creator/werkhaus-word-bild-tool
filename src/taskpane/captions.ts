import type { ImageItem } from "./types";

// Stabiler Kernbereich:
// Caption-Regeln und Caption-Entscheidung bleiben getrennt von Import und Word-Ausgabe.
// Nur bei fachlicher Regelanderung anpassen.

export function getResolvedCaptionText(item: ImageItem, figureNumber: number): string {
  const trimmedCaption = item.caption?.trim() || "";
  const trimmedName = item.name?.trim() || "";

  if (trimmedCaption.length > 0) {
    return `Abbildung ${figureNumber}: ${trimmedCaption}`;
  }

  if (trimmedName.length > 0) {
    return `Abbildung ${figureNumber}: ${trimmedName}`;
  }

  return `Abbildung ${figureNumber}: Bild ohne Bezeichnung`;
}

export function shouldInsertCaption(item: ImageItem): boolean {
  return item.includeCaptionInWord !== false;
}

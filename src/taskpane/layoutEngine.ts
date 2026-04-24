import { CONTACT_SHEET_AREA_MM, MULTI_INSERT_LAYOUTS } from "./constants";
import { shouldInsertCaption } from "./captions";
import type {
  CellPlan,
  ImageItem,
  LayoutKind,
  LayoutPlan,
  LayoutSpec,
  MultiInsertOptions,
  PagePlan,
  SizeMm,
} from "./types";

const DEFAULT_INSET_MM = 2;

// Reine Fachlogik ohne DOM und ohne Word-API.
// Diese Schicht berechnet nur Geometrie und Planungsgrundlagen fuer Mehrbild-Einfuellungen.

export function resolveLayoutGrid(layoutKind: LayoutKind): { columns: number; rows: number } {
  const definition = MULTI_INSERT_LAYOUTS[layoutKind];

  if (!definition) {
    throw new Error(`Unsupported layout kind: ${layoutKind}`);
  }

  return definition;
}

export function createLayoutSpec(options: MultiInsertOptions): LayoutSpec {
  const { columns, rows } = resolveLayoutGrid(options.layoutKind);
  const contentWidthMm = Math.max(0, options.pageWidthMm ?? CONTACT_SHEET_AREA_MM.width);
  const contentHeightMm = Math.max(0, options.pageHeightMm ?? CONTACT_SHEET_AREA_MM.height);
  const insetMm = Math.max(0, options.insetMm ?? DEFAULT_INSET_MM);
  const captionHeightMm = Math.max(0, options.captionHeightMm ?? 0);
  const itemsPerPage = columns * rows;
  const cellWidthMm = contentWidthMm / columns;
  const cellHeightMm = contentHeightMm / rows;
  const imageBoxWidthMm = Math.max(0, cellWidthMm - insetMm * 2);
  const imageBoxHeightMm = Math.max(0, cellHeightMm - insetMm * 2 - captionHeightMm);

  return {
    layoutKind: options.layoutKind,
    columns,
    rows,
    itemsPerPage,
    contentWidthMm,
    contentHeightMm,
    insetMm,
    captionHeightMm,
    cellWidthMm,
    cellHeightMm,
    imageBoxWidthMm,
    imageBoxHeightMm,
  };
}

export function fitImageIntoBox(
  sourceWidthMm: number,
  sourceHeightMm: number,
  boxWidthMm: number,
  boxHeightMm: number
): SizeMm | undefined {
  if (
    !Number.isFinite(sourceWidthMm) ||
    !Number.isFinite(sourceHeightMm) ||
    !Number.isFinite(boxWidthMm) ||
    !Number.isFinite(boxHeightMm) ||
    sourceWidthMm <= 0 ||
    sourceHeightMm <= 0 ||
    boxWidthMm <= 0 ||
    boxHeightMm <= 0
  ) {
    return undefined;
  }

  const scale = Math.min(boxWidthMm / sourceWidthMm, boxHeightMm / sourceHeightMm);

  if (!Number.isFinite(scale) || scale <= 0) {
    return undefined;
  }

  return {
    widthMm: roundMm(sourceWidthMm * scale),
    heightMm: roundMm(sourceHeightMm * scale),
  };
}

export function createCellPlan(
  item: ImageItem | undefined,
  pageIndex: number,
  cellIndex: number,
  rowIndex: number,
  columnIndex: number,
  spec: LayoutSpec
): CellPlan {
  const isEmpty = !item;
  const captionEnabled = !!item && spec.captionHeightMm > 0 && shouldInsertCaption(item);
  const captionText = item ? buildCaptionText(item) : undefined;
  const fittedImageSizeMm = item
    ? fitImageIntoBox(
        item.exif?.width || 0,
        item.exif?.height || 0,
        spec.imageBoxWidthMm,
        spec.imageBoxHeightMm
      )
    : undefined;

  return {
    pageIndex,
    cellIndex,
    rowIndex,
    columnIndex,
    item,
    isEmpty,
    cellWidthMm: spec.cellWidthMm,
    cellHeightMm: spec.cellHeightMm,
    imageBoxWidthMm: spec.imageBoxWidthMm,
    imageBoxHeightMm: spec.imageBoxHeightMm,
    captionHeightMm: spec.captionHeightMm,
    fittedImageSizeMm,
    captionText,
    captionEnabled,
  };
}

export function createPagePlan(items: ImageItem[], pageIndex: number, spec: LayoutSpec): PagePlan {
  const cells: CellPlan[] = [];

  for (let cellIndex = 0; cellIndex < spec.itemsPerPage; cellIndex += 1) {
    const rowIndex = Math.floor(cellIndex / spec.columns);
    const columnIndex = cellIndex % spec.columns;
    const item = items[cellIndex];

    cells.push(createCellPlan(item, pageIndex, cellIndex, rowIndex, columnIndex, spec));
  }

  return {
    pageIndex,
    layoutKind: spec.layoutKind,
    spec,
    cells,
    imageCount: items.length,
  };
}

export function createLayoutPlan(pages: PagePlan[], spec: LayoutSpec): LayoutPlan {
  return {
    spec,
    pages,
    totalImageCount: pages.reduce((sum, page) => sum + page.imageCount, 0),
  };
}

function buildCaptionText(item: ImageItem): string {
  const trimmedCaption = item.caption?.trim() || "";
  const trimmedName = item.name?.trim() || "";

  if (trimmedCaption.length > 0) {
    return trimmedCaption;
  }

  if (trimmedName.length > 0) {
    return trimmedName;
  }

  return "";
}

function roundMm(value: number): number {
  return Number(value.toFixed(2));
}

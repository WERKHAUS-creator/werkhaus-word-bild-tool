import type { LayoutKind, PageConfig } from "./types";

export const PAGE_CONFIG: PageConfig = {
  pageWidthMm: 210,
  pageHeightMm: 297,
  marginTopMm: 20,
  marginBottomMm: 20,
  marginLeftMm: 20,
  marginRightMm: 20,
};

export const MM_PER_CM = 10;

export const CONTACT_SHEET_AREA_CM = {
  width: 16,
  height: 23,
} as const;

export const CONTACT_SHEET_AREA_MM = {
  width: CONTACT_SHEET_AREA_CM.width * MM_PER_CM,
  height: CONTACT_SHEET_AREA_CM.height * MM_PER_CM,
} as const;

export const MULTI_INSERT_LAYOUTS: Record<LayoutKind, { columns: number; rows: number }> = {
  1: { columns: 1, rows: 1 },
  2: { columns: 1, rows: 2 },
  4: { columns: 2, rows: 2 },
  6: { columns: 2, rows: 3 },
  8: { columns: 2, rows: 4 },
  12: { columns: 3, rows: 4 },
  16: { columns: 4, rows: 4 },
};

export const USABLE_PAGE_WIDTH_CM =
  (PAGE_CONFIG.pageWidthMm - PAGE_CONFIG.marginLeftMm - PAGE_CONFIG.marginRightMm) / MM_PER_CM;
export const USABLE_PAGE_HEIGHT_CM =
  (PAGE_CONFIG.pageHeightMm - PAGE_CONFIG.marginTopMm - PAGE_CONFIG.marginBottomMm) / MM_PER_CM;

export const WORD_PT_PER_INCH = 72;
export const MM_PER_INCH = 25.4;

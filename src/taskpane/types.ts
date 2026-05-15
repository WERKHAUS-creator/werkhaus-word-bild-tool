export interface ImageItem {
  id: string;
  key: string;
  name: string;
  relativePath?: string;
  size: number;
  lastModified?: number;
  base64: string;
  previewUrl: string;
  hash: string;
  caption: string;
  position: number;
  selected: boolean;
  includeCaptionInWord: boolean;
  exif?: {
    dateTimeOriginal?: string;
    model?: string;
    width?: number;
    height?: number;
  };
}

export type CaptionMode = "full" | "numberOnly" | "plainImage";
export type SortMode = "exifDate" | "name" | "custom";

export type LayoutKind = 1 | 2 | 4 | 6 | 8 | 12 | 16;

export interface SizeMm {
  widthMm: number;
  heightMm: number;
}

export interface LayoutSpec {
  layoutKind: LayoutKind;
  columns: number;
  rows: number;
  itemsPerPage: number;
  contentWidthMm: number;
  contentHeightMm: number;
  insetMm: number;
  captionHeightMm: number;
  cellWidthMm: number;
  cellHeightMm: number;
  imageBoxWidthMm: number;
  imageBoxHeightMm: number;
}

export interface CellPlan {
  pageIndex: number;
  cellIndex: number;
  rowIndex: number;
  columnIndex: number;
  item?: ImageItem;
  isEmpty: boolean;
  cellWidthMm: number;
  cellHeightMm: number;
  imageBoxWidthMm: number;
  imageBoxHeightMm: number;
  captionHeightMm: number;
  fittedImageSizeMm?: SizeMm;
  captionText?: string;
  captionEnabled: boolean;
}

export interface PagePlan {
  pageIndex: number;
  layoutKind: LayoutKind;
  spec: LayoutSpec;
  cells: CellPlan[];
  imageCount: number;
}

export interface LayoutPlan {
  spec: LayoutSpec;
  pages: PagePlan[];
  totalImageCount: number;
}

export interface MultiInsertOptions {
  layoutKind: LayoutKind;
  pageWidthMm?: number;
  pageHeightMm?: number;
  insetMm?: number;
  captionHeightMm?: number;
}

export interface StoredMeta {
  caption?: string;
  comment?: string;
  importedAt?: string;
  position?: number;
  selected?: boolean;
  includeCaptionInWord?: boolean;
}

export interface PageConfig {
  pageWidthMm: number;
  pageHeightMm: number;
  marginTopMm: number;
  marginBottomMm: number;
  marginLeftMm: number;
  marginRightMm: number;
}

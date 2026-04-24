import { createLayoutPlan, createLayoutSpec, createPagePlan } from "./layoutEngine";
import type { ImageItem, LayoutPlan, MultiInsertOptions, PagePlan } from "./types";

// Reine Verteilungsschicht:
// Sortierte Bilder werden ohne Word-API in Seitenpakete aufgeteilt.

export function paginateImages(items: ImageItem[], options: MultiInsertOptions): LayoutPlan {
  const spec = createLayoutSpec(options);
  const pages: PagePlan[] = [];

  for (let index = 0; index < items.length; index += spec.itemsPerPage) {
    const pageItems = items.slice(index, index + spec.itemsPerPage);
    pages.push(createPagePlan(pageItems, pages.length, spec));
  }

  return createLayoutPlan(pages, spec);
}

export function splitIntoPages(items: ImageItem[], itemsPerPage: number): ImageItem[][] {
  const safeItemsPerPage = Math.max(1, Math.floor(itemsPerPage));
  const pages: ImageItem[][] = [];

  for (let index = 0; index < items.length; index += safeItemsPerPage) {
    pages.push(items.slice(index, index + safeItemsPerPage));
  }

  return pages;
}

import type { CaptionMode, ImageItem, SortMode } from "./types";
import { IMAGE_FILE_INPUT_ACCEPT, importImageFiles } from "./importImages";
import { createTaskpanePersistence } from "./persistence";
import {
  type BilddatenProjectFile,
  type BilddatenProjectImage,
  PROJECT_FILE_NAME,
  buildBilddatenProjectFile,
  matchProjectImagesToItems,
  parseBilddatenProjectFile,
  serializeBilddatenProjectFile,
} from "./projectFile";
import { insertSingleImageAtSelection } from "./wordInsert";

type ProjectFolderHandle = {
  getFileHandle(
    name: string,
    options?: { create?: boolean }
  ): Promise<{
    getFile(): Promise<File>;
    createWritable(): Promise<{
      write(data: string): Promise<void>;
      close(): Promise<void>;
    }>;
  }>;
  values(): AsyncIterableIterator<any>;
};

// Stabiler Kernbereich:
// Bildimport, Bildauswahl und Einzelfoto-Einfuegen bleiben hier bewusst getrennt.
// Nur bei nachgewiesenem Fehler an diesen Pfaden aendern.
let imageItems: ImageItem[] = [];
let currentFolderHandle: ProjectFolderHandle | null = null;
let currentProjectCreatedAt = new Date().toISOString();
const DEFAULT_PREVIEW_SIZE_PX = 120;
const MIN_PREVIEW_SIZE_PX = 120;
const MAX_PREVIEW_SIZE_PX = 500;
const MAX_INSERT_SIZE_CM = 16;
const naturalSortCollator = new Intl.Collator("de", { numeric: true, sensitivity: "base" });

function initTaskpane() {
  const {
    loadAllMeta,
    loadSettings,
    getMeta,
    setMeta,
    getSortMode,
    setSortMode,
    getCollapsedSections,
    setCollapsedSections,
  } = createTaskpanePersistence();
  const statusElement = document.getElementById("statusMessage");
  const imageList = document.getElementById("imageList");
  const dropZone = document.getElementById("imageDropZone");

  const imageUpload = document.getElementById("imageUpload") as HTMLInputElement | null;
  const folderUpload = document.getElementById("folderUpload") as HTMLInputElement | null;

  const pickFilesButton = document.getElementById("pickFilesButton");
  const pickFolderButton = document.getElementById("pickFolderButton");
  const saveProjectButton = document.getElementById("saveProjectButton");
  const loadProjectButton = document.getElementById("loadProjectButton");
  const expandAllSectionsButton = document.getElementById("expandAllSectionsButton");
  const collapseAllSectionsButton = document.getElementById("collapseAllSectionsButton");
  const clearImagesButton = document.getElementById("clearImagesButton");
  const toggleInfoButton = document.getElementById("toggleInfoButton");
  const toggleCaptionButton = document.getElementById("toggleCaptionButton");
  const sortModeSelect = document.getElementById("sortModeSelect") as HTMLSelectElement | null;
  const selectAllButton = document.getElementById("selectAllButton");
  const selectNoneButton = document.getElementById("selectNoneButton");
  const insertSelectedButton = document.getElementById("insertSelectedButton");
  const insertSelectedCaptionButton = document.getElementById("insertSelectedCaptionButton");
  const insertSelectedPlainButton = document.getElementById("insertSelectedPlainButton");

  const previewSizeRange = document.getElementById("previewSizeRange") as HTMLInputElement | null;
  const previewSizeValue = document.getElementById("previewSizeValue");
  const insertSizeRange = document.getElementById("insertSizeRange") as HTMLInputElement | null;
  const insertSizeValue = document.getElementById("insertSizeValue");
  const projectImportUpload = document.getElementById(
    "projectImportUpload"
  ) as HTMLInputElement | null;

  loadAllMeta();
  loadSettings();

  let insertSizeCm = 10;
  let previewSizePx = DEFAULT_PREVIEW_SIZE_PX;
  let sortMode: SortMode = getSortMode() || "custom";

  let showInfo = false;
  let showCaptions = false;
  let dropZoneDragDepth = 0;

  if (previewSizeRange) {
    previewSizeRange.value = String(DEFAULT_PREVIEW_SIZE_PX);
    previewSizeRange.addEventListener("input", () => {
      const nextSize = Math.min(
        MAX_PREVIEW_SIZE_PX,
        Math.max(MIN_PREVIEW_SIZE_PX, Number(previewSizeRange.value) || DEFAULT_PREVIEW_SIZE_PX)
      );
      previewSizePx = nextSize;
      previewSizeRange.value = String(nextSize);
      updatePreviewSize();
    });
  }

  if (insertSizeRange && insertSizeValue) {
    insertSizeCm = Math.min(MAX_INSERT_SIZE_CM, Number(insertSizeRange.value));
    insertSizeValue.textContent = `${insertSizeCm.toFixed(1)} cm`;

    insertSizeRange.addEventListener("input", () => {
      insertSizeCm = Math.min(MAX_INSERT_SIZE_CM, Number(insertSizeRange.value));
      insertSizeValue.textContent = `${insertSizeCm.toFixed(1)} cm`;
    });
  }

  if (imageUpload) {
    imageUpload.accept = IMAGE_FILE_INPUT_ACCEPT;
  }

  if (folderUpload) {
    folderUpload.accept = IMAGE_FILE_INPUT_ACCEPT;
  }

  function setStatus(message: string) {
    if (statusElement) {
      statusElement.textContent = message;
    }
  }

  function setDropZoneActive(active: boolean) {
    dropZone?.classList.toggle("drop-zone-active", active);
  }

  function resetDropZoneState() {
    dropZoneDragDepth = 0;
    setDropZoneActive(false);
  }

  function hasFileTransfer(dataTransfer: DataTransfer | null): boolean {
    if (!dataTransfer) {
      return false;
    }

    return Array.from(dataTransfer.types || []).includes("Files");
  }

  function getSupportedFormatLabel(): string {
    return IMAGE_FILE_INPUT_ACCEPT.replace(/,/g, ", ").toUpperCase();
  }

  function buildNoValidImageStatus(
    importResult: Awaited<ReturnType<typeof importImageFiles>>
  ): string {
    if (importResult.invalidFileCount === 0) {
      return "Keine gültigen Bilddateien gefunden.";
    }

    const invalidPreview = importResult.invalidFileNames.slice(0, 3).join(", ");
    const invalidSuffix = importResult.invalidFileNames.length > 3 ? ", ..." : "";

    return `Keine Bilder übernommen. Es wurden nur nicht unterstützte Dateien erkannt: ${invalidPreview}${invalidSuffix}. Unterstützt: ${getSupportedFormatLabel()}.`;
  }

  function buildImportStatus(importResult: Awaited<ReturnType<typeof importImageFiles>>): string {
    const parts: string[] = [];

    if (importResult.addedCount > 0) {
      parts.push(`${importResult.addedCount} Bild(er) hinzugefügt.`);
    } else {
      parts.push("Keine neuen Bilder hinzugefügt.");
    }

    if (importResult.skippedCount > 0) {
      parts.push(`${importResult.skippedCount} Dublette(n) übersprungen.`);
    }

    if (importResult.invalidFileCount > 0) {
      const invalidPreview = importResult.invalidFileNames.slice(0, 3).join(", ");
      const invalidSuffix = importResult.invalidFileNames.length > 3 ? ", ..." : "";
      parts.push(
        `${importResult.invalidFileCount} Datei(en) ignoriert: ${invalidPreview}${invalidSuffix}. Unterstützt: ${getSupportedFormatLabel()}.`
      );
    }

    if (importResult.failedCount > 0) {
      const failedPreview = importResult.failedFileNames.slice(0, 3).join(", ");
      const failedSuffix = importResult.failedFileNames.length > 3 ? ", ..." : "";
      parts.push(
        `${importResult.failedCount} Datei(en) konnten nicht gelesen werden: ${failedPreview}${failedSuffix}.`
      );
    }

    return parts.join(" ");
  }

  function formatFileSize(bytes: number): string {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function updatePreviewSize() {
    document.documentElement.style.setProperty("--preview-size", `${previewSizePx}px`);
    if (previewSizeValue) {
      previewSizeValue.textContent = `${previewSizePx} px`;
    }
  }

  function getSortModeLabel(mode: SortMode): string {
    if (mode === "exifDate") {
      return "EXIF-Aufnahmedatum";
    }

    if (mode === "name") {
      return "Bildname";
    }

    return "Eigene Sortierung";
  }

  function updateSortUI() {
    if (sortModeSelect && sortModeSelect.value !== sortMode) {
      sortModeSelect.value = sortMode;
    }

    if (sortModeSelect) {
      const sortLabel = getSortModeLabel(sortMode);
      sortModeSelect.setAttribute("aria-label", `Sortierung: ${sortLabel}`);
      sortModeSelect.setAttribute("title", `Aktuell: ${sortLabel}`);
    }
  }

  function getCollapsibleSections(): HTMLElement[] {
    return Array.from(document.querySelectorAll<HTMLElement>(".collapsible-section"));
  }

  function getSectionKey(section: HTMLElement): string | null {
    const sectionKey = section.getAttribute("data-section-key");
    return sectionKey && sectionKey.length > 0 ? sectionKey : null;
  }

  function updateSectionToggleButton(section: HTMLElement, collapsed: boolean) {
    const toggleButton = section.querySelector<HTMLButtonElement>("[data-section-toggle]");
    if (!toggleButton) {
      return;
    }

    toggleButton.textContent = "";
    toggleButton.setAttribute("aria-expanded", String(!collapsed));
    toggleButton.setAttribute(
      "aria-label",
      collapsed ? "Bereich ausklappen" : "Bereich einklappen"
    );
    toggleButton.setAttribute("title", collapsed ? "Bereich ausklappen" : "Bereich einklappen");
  }

  function setSectionCollapsed(sectionKey: string, collapsed: boolean, persistState = true) {
    const section = document.querySelector<HTMLElement>(`[data-section-key="${sectionKey}"]`);
    if (!section) {
      return;
    }

    section.classList.toggle("is-collapsed", collapsed);
    updateSectionToggleButton(section, collapsed);

    if (persistState) {
      persistCollapsedSectionsState();
    }
  }

  function persistCollapsedSectionsState() {
    const collapsedSectionKeys = getCollapsibleSections()
      .filter((section) => section.classList.contains("is-collapsed"))
      .map((section) => getSectionKey(section))
      .filter((sectionKey): sectionKey is string => sectionKey !== null);

    setCollapsedSections(collapsedSectionKeys);
  }

  function setAllSectionsCollapsed(collapsed: boolean, persistState = true) {
    for (const section of getCollapsibleSections()) {
      section.classList.toggle("is-collapsed", collapsed);
      updateSectionToggleButton(section, collapsed);
    }

    if (persistState) {
      persistCollapsedSectionsState();
    }
  }

  function restoreCollapsedSections() {
    const storedCollapsedSections = new Set(getCollapsedSections());

    setAllSectionsCollapsed(false, false);

    for (const section of getCollapsibleSections()) {
      const sectionKey = getSectionKey(section);
      if (!sectionKey) {
        continue;
      }

      if (storedCollapsedSections.has(sectionKey)) {
        setSectionCollapsed(sectionKey, true, false);
      }
    }
  }

  function bindSectionToggleButtons() {
    const toggleButtons = Array.from(
      document.querySelectorAll<HTMLButtonElement>("[data-section-toggle]")
    );

    toggleButtons.forEach((button) => {
      button.addEventListener("click", () => {
        const sectionKey = button.getAttribute("data-section-toggle");
        if (!sectionKey) {
          return;
        }

        const section = document.querySelector<HTMLElement>(`[data-section-key="${sectionKey}"]`);
        if (!section) {
          return;
        }

        const collapsed = !section.classList.contains("is-collapsed");
        setSectionCollapsed(sectionKey, collapsed);
      });
    });
  }

  function parseExifTimestamp(rawValue?: string): number | null {
    if (!rawValue) {
      return null;
    }

    const normalized = rawValue.trim();
    const match = normalized.match(/^(\d{4}):(\d{2}):(\d{2})(?:[ T](\d{2}):(\d{2}):(\d{2}))?$/);

    if (!match) {
      return null;
    }

    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const hour = Number(match[4] || "0");
    const minute = Number(match[5] || "0");
    const second = Number(match[6] || "0");

    const candidate = new Date(year, month - 1, day, hour, minute, second);
    if (
      candidate.getFullYear() !== year ||
      candidate.getMonth() !== month - 1 ||
      candidate.getDate() !== day ||
      candidate.getHours() !== hour ||
      candidate.getMinutes() !== minute ||
      candidate.getSeconds() !== second
    ) {
      return null;
    }

    const timestamp = candidate.getTime();
    return Number.isFinite(timestamp) ? timestamp : null;
  }

  function compareByNaturalName(left: ImageItem, right: ImageItem): number {
    return naturalSortCollator.compare(left.name, right.name);
  }

  function compareByExifDate(left: ImageItem, right: ImageItem): number {
    const leftTimestamp = parseExifTimestamp(left.exif?.dateTimeOriginal);
    const rightTimestamp = parseExifTimestamp(right.exif?.dateTimeOriginal);

    if (leftTimestamp !== null && rightTimestamp !== null) {
      if (leftTimestamp !== rightTimestamp) {
        return leftTimestamp - rightTimestamp;
      }

      return compareByNaturalName(left, right);
    }

    return compareByNaturalName(left, right);
  }

  function sortItemsForCurrentMode(): void {
    if (sortMode === "exifDate") {
      imageItems.sort(compareByExifDate);
    } else if (sortMode === "name") {
      imageItems.sort(compareByNaturalName);
    }
  }

  function setCurrentSortMode(nextMode: SortMode) {
    sortMode = nextMode;
    setSortMode(nextMode);
    updateSortUI();
  }

  function applySortModeToItems(options: { render?: boolean } = {}) {
    sortItemsForCurrentMode();
    persistImagePositions();

    if (options.render !== false) {
      renderImageList();
    }

    updateSortUI();
  }

  function applyProjectImageToItem(item: ImageItem, projectImage: BilddatenProjectImage): void {
    item.relativePath = projectImage.relativePath || item.relativePath;
    item.lastModified = projectImage.lastModified ?? item.lastModified;
    item.caption = projectImage.caption || "";
    item.selected = projectImage.active;
    item.includeCaptionInWord = projectImage.includeCaptionInWord;
    item.position = projectImage.position;

    if (projectImage.exifDateTaken) {
      item.exif = {
        ...(item.exif || {}),
        dateTimeOriginal: projectImage.exifDateTaken,
      };
    }

    setMeta(item.hash, {
      caption: item.caption,
      position: projectImage.position,
      selected: projectImage.active,
      includeCaptionInWord: projectImage.includeCaptionInWord,
    });
  }

  function applyProjectFileToLoadedImages(projectFile: BilddatenProjectFile): {
    applied: boolean;
    matchedCount: number;
    unmatchedProjectCount: number;
    unmatchedItemCount: number;
  } {
    currentProjectCreatedAt = projectFile.createdAt || currentProjectCreatedAt;
    setCurrentSortMode(projectFile.sortMode || "custom");

    if (imageItems.length === 0) {
      return {
        applied: false,
        matchedCount: 0,
        unmatchedProjectCount: projectFile.images.length,
        unmatchedItemCount: 0,
      };
    }

    const matchResult = matchProjectImagesToItems(projectFile, imageItems);

    if (matchResult.matches.length === 0) {
      return {
        applied: false,
        matchedCount: 0,
        unmatchedProjectCount: projectFile.images.length,
        unmatchedItemCount: imageItems.length,
      };
    }

    const orderedMatches = [...matchResult.matches].sort((left, right) => {
      const positionDiff = left.image.position - right.image.position;
      if (positionDiff !== 0) {
        return positionDiff;
      }

      return left.itemIndex - right.itemIndex;
    });

    const reorderedItems: ImageItem[] = [];

    for (const match of orderedMatches) {
      const item = imageItems[match.itemIndex];
      applyProjectImageToItem(item, match.image);
      reorderedItems.push(item);
    }

    for (const itemIndex of matchResult.unmatchedItemIndexes) {
      reorderedItems.push(imageItems[itemIndex]);
    }

    imageItems = reorderedItems;
    persistImagePositions();
    renderImageList();

    return {
      applied: true,
      matchedCount: matchResult.matches.length,
      unmatchedProjectCount: matchResult.unmatchedProjectImages.length,
      unmatchedItemCount: matchResult.unmatchedItemIndexes.length,
    };
  }

  function buildProjectSummaryMessage(
    prefix: string,
    result: {
      applied: boolean;
      matchedCount: number;
      unmatchedProjectCount: number;
      unmatchedItemCount: number;
    }
  ): string {
    if (!result.applied) {
      if (result.unmatchedProjectCount > 0) {
        return `${prefix}: Keine passenden Bilder gefunden.`;
      }

      return `${prefix}: Keine Bilder geladen.`;
    }

    const parts = [`${prefix}: ${result.matchedCount} Bild(er) angewendet.`];

    if (result.unmatchedProjectCount > 0) {
      parts.push(`${result.unmatchedProjectCount} Eintrag(e) ohne Treffer.`);
    }

    if (result.unmatchedItemCount > 0) {
      parts.push(`${result.unmatchedItemCount} Bild(er) ohne Projektzuordnung.`);
    }

    return parts.join(" ");
  }

  async function readProjectFileText(file: File): Promise<string> {
    if (typeof file.text === "function") {
      return file.text();
    }

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        if (typeof reader.result === "string") {
          resolve(reader.result);
          return;
        }

        reject(new Error("Projektdatei konnte nicht gelesen werden."));
      };
      reader.onerror = () => reject(new Error("Projektdatei konnte nicht gelesen werden."));
      reader.readAsText(file);
    });
  }

  async function readProjectFileFromFolderHandle(
    folderHandle: ProjectFolderHandle
  ): Promise<BilddatenProjectFile | undefined> {
    try {
      const fileHandle = await folderHandle.getFileHandle(PROJECT_FILE_NAME, { create: false });
      const file = await fileHandle.getFile();
      const rawText = await file.text();
      return parseBilddatenProjectFile(rawText);
    } catch {
      return undefined;
    }
  }

  async function autoLoadProjectFileFromCurrentFolder(): Promise<string | undefined> {
    if (!currentFolderHandle) {
      return undefined;
    }

    const projectFile = await readProjectFileFromFolderHandle(currentFolderHandle);
    if (!projectFile) {
      return undefined;
    }

    const result = applyProjectFileToLoadedImages(projectFile);
    return buildProjectSummaryMessage(`bilddaten.json automatisch geladen`, result);
  }

  async function writeProjectFileToCurrentFolder(
    serializedProjectFile: string
  ): Promise<"saved" | "cancelled" | "failed"> {
    if (!currentFolderHandle) {
      return "failed";
    }

    try {
      let existingFile = false;
      try {
        await currentFolderHandle.getFileHandle(PROJECT_FILE_NAME, { create: false });
        existingFile = true;
      } catch {
        existingFile = false;
      }

      if (existingFile) {
        const overwrite = window.confirm(
          "bilddaten.json existiert bereits. Soll die Datei überschrieben werden?"
        );
        if (!overwrite) {
          setStatus("Speichern abgebrochen.");
          return "cancelled";
        }
      }

      const fileHandle = await currentFolderHandle.getFileHandle(PROJECT_FILE_NAME, {
        create: true,
      });
      const writable = await fileHandle.createWritable();
      await writable.write(serializedProjectFile);
      await writable.close();
      return "saved";
    } catch (error) {
      console.error("Fehler beim Speichern von bilddaten.json:", error);
      return "failed";
    }
  }

  async function saveProjectFile(): Promise<void> {
    if (imageItems.length === 0) {
      setStatus("Keine Bilder zum Speichern vorhanden.");
      return;
    }

    const projectFile = buildBilddatenProjectFile(imageItems, sortMode, currentProjectCreatedAt);
    const serializedProjectFile = serializeBilddatenProjectFile(projectFile);

    if (currentFolderHandle) {
      const saveResult = await writeProjectFileToCurrentFolder(serializedProjectFile);
      if (saveResult === "saved" || saveResult === "cancelled") {
        currentProjectCreatedAt = projectFile.createdAt;
        if (saveResult === "saved") {
          setStatus(`bilddaten.json wurde im aktuellen Ordner gespeichert.`);
        }
        return;
      }
    }

    const savePicker = (window as any).showSaveFilePicker;
    if (typeof savePicker === "function") {
      try {
        const fileHandle = await savePicker({
          suggestedName: PROJECT_FILE_NAME,
          types: [
            {
              description: "WERKHAUS Bilddaten",
              accept: {
                "application/json": [".json"],
              },
            },
          ],
        });

        const writable = await fileHandle.createWritable();
        await writable.write(serializedProjectFile);
        await writable.close();
        currentProjectCreatedAt = projectFile.createdAt;
        setStatus(`bilddaten.json wurde gespeichert.`);
        return;
      } catch (error) {
        console.warn("Speichern über showSaveFilePicker fehlgeschlagen oder abgebrochen:", error);
        setStatus("Speichern abgebrochen.");
        return;
      }
    }

    setStatus("Speichern nicht unterstützt. Bitte Ordner über die native Ordnerauswahl laden.");
  }

  async function loadProjectFromText(rawText: string, sourceLabel: string): Promise<void> {
    const projectFile = parseBilddatenProjectFile(rawText);
    if (!projectFile) {
      setStatus(`${sourceLabel}: bilddaten.json konnte nicht gelesen werden.`);
      return;
    }

    const result = applyProjectFileToLoadedImages(projectFile);
    const summary = buildProjectSummaryMessage(sourceLabel, result);
    setStatus(summary);
  }

  async function loadProjectFromSelectedFile(file: File, sourceLabel: string): Promise<void> {
    try {
      const rawText = await readProjectFileText(file);
      await loadProjectFromText(rawText, sourceLabel);
    } catch (error) {
      console.error("Fehler beim Laden der Projektdatei:", error);
      setStatus(`${sourceLabel}: bilddaten.json konnte nicht gelesen werden.`);
    }
  }

  function resetInputValue(input: HTMLInputElement | null, label: string) {
    try {
      if (input) input.value = "";
    } catch (error) {
      console.warn(`Konnte ${label} nicht zuruecksetzen:`, error);
    }
  }

  function persistImagePositions() {
    imageItems.forEach((item, index) => {
      item.position = index + 1;
      setMeta(item.hash, { position: item.position });
    });
  }

  function clearImageList() {
    try {
      for (const item of imageItems) {
        if (item.previewUrl && item.previewUrl.startsWith("blob:")) {
          try {
            URL.revokeObjectURL(item.previewUrl);
          } catch {
            // Blob-URLs muessen nur bei echten Blob-Quellen freigegeben werden; andere Quellen sind hier unkritisch.
          }
        }
      }
    } catch {
      // Ein fehlgeschlagenes Freigeben einzelner Vorschauquellen darf das Leeren der Liste nicht blockieren.
    }

    imageItems = [];

    resetInputValue(imageUpload, "imageUpload");
    resetInputValue(folderUpload, "folderUpload");

    renderImageList();
    setStatus("Bildliste geleert.");
  }

  function updateVisibilityUI() {
    if (imageList) {
      imageList.classList.toggle("hide-info", !showInfo);
      imageList.classList.toggle("hide-captions", !showCaptions);
    }

    if (toggleInfoButton) {
      toggleInfoButton.textContent = "Bild Infos";
    }

    if (toggleCaptionButton) {
      toggleCaptionButton.textContent = "Beschriftung";
    }
  }

  function getSelectedItems(): ImageItem[] {
    return imageItems.filter((item) => item.selected !== false);
  }

  function setAllItemsSelected(selected: boolean) {
    imageItems.forEach((item) => {
      item.selected = selected;
      setMeta(item.hash, { selected });
    });

    renderImageList();
    setStatus(selected ? "Alle Bilder sind ausgewählt." : "Alle Bilder sind abgewählt.");
  }

  async function insertSpacerParagraph() {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const spacer = selection.insertParagraph("", Word.InsertLocation.after);
      spacer.getRange(Word.RangeLocation.end).select();
      await context.sync();
    });
  }

  async function insertSelectedImages(captionMode: CaptionMode) {
    const selectedItems = getSelectedItems();

    if (selectedItems.length === 0) {
      setStatus("Keine Bilder ausgewählt.");
      return;
    }

    try {
      if (typeof (window as any).Word === "undefined") {
        throw new Error("Word API is not available in this host (Word is undefined)");
      }

      for (const [index, item] of selectedItems.entries()) {
        await insertSingleImageAtSelection(item, insertSizeCm, {
          includeCaption: true,
          captionMode,
        });

        if (captionMode === "numberOnly" && index < selectedItems.length - 1) {
          await insertSpacerParagraph();
        }
      }

      setStatus(
        captionMode === "full"
          ? `${selectedItems.length} ausgewählte Bilder wurden mit Beschriftung in Word eingefügt.`
          : captionMode === "numberOnly"
            ? `${selectedItems.length} ausgewählte Bilder wurden nur mit Nummerierung in Word eingefügt.`
            : `${selectedItems.length} ausgewählte Bilder wurden ohne Beschriftung in Word eingefügt.`
      );
    } catch (e: any) {
      console.error("Fehler beim Einfügen ausgewählter Bilder:", e);
      console.error("DebugInfo:", e?.debugInfo);
      setStatus(`Fehler beim Word.run: ${e?.message || "Unbekannter Fehler"}`);
    }
  }

  async function appendFiles(files: FileList | File[]) {
    // Stabiler Kernbereich: Datei-, Ordner- und Drop-Import nutzen denselben Importpfad.
    const importResult = await importImageFiles(files, imageItems, { getMeta, setMeta });

    if (importResult.imageFileCount === 0) {
      setStatus(buildNoValidImageStatus(importResult));
      return;
    }

    try {
      imageItems = importResult.items;
      applySortModeToItems({ render: false });

      resetInputValue(imageUpload, "imageUpload");
      resetInputValue(folderUpload, "folderUpload");

      renderImageList();
      setStatus(buildImportStatus(importResult));
    } catch (error) {
      console.error(error);
      setStatus("Fehler beim Laden der Bilder.");
    }
  }

  function isInteractiveElement(target: EventTarget | null): boolean {
    if (!(target instanceof HTMLElement)) return false;
    return !!target.closest("button, input, textarea, select, label");
  }

  function renderImageList() {
    // UI-Orchestrierung: nur Listenansicht und Bedienung, nicht mit Word-Ausgabe vermischen.
    if (!imageList) return;

    imageList.classList.toggle("hide-info", !showInfo);
    imageList.classList.toggle("hide-captions", !showCaptions);
    imageList.innerHTML = "";

    if (imageItems.length === 0) {
      const emptyState = document.createElement("div");
      emptyState.className = "empty-state";
      emptyState.textContent = "Noch keine Bilder ausgewählt.";
      imageList.appendChild(emptyState);
      return;
    }

    const itemsToRender = [...imageItems];

    itemsToRender.forEach((item, index) => {
      const card = document.createElement("div");
      card.className = "image-card";
      card.draggable = true;
      card.setAttribute("data-id", item.id);

      card.addEventListener("dragstart", (ev) => {
        if (isInteractiveElement(ev.target)) {
          ev.preventDefault();
          return;
        }

        try {
          (ev.dataTransfer as DataTransfer).setData("text/plain", item.id);
        } catch {
          // Ohne Drag-Payload bleibt nur das visuelle Dragging aktiv; die UI bleibt trotzdem benutzbar.
        }
        card.classList.add("dragging");
      });

      card.addEventListener("dragend", () => {
        card.classList.remove("dragging");
        document.querySelectorAll(".image-card.drag-over").forEach((el) => {
          el.classList.remove("drag-over");
        });
      });

      card.addEventListener("dragover", (ev) => {
        if (hasFileTransfer(ev.dataTransfer)) return;
        if (isInteractiveElement(ev.target)) return;

        ev.preventDefault();
        const target = ev.currentTarget as HTMLElement;
        target.classList.add("drag-over");
        try {
          (ev.dataTransfer as DataTransfer).dropEffect = "move";
        } catch {
          // Manche Hosts erlauben das Setzen des Drop-Effekts nicht; das Drag-and-drop soll trotzdem weiterlaufen.
        }
      });

      card.addEventListener("dragleave", (ev) => {
        if (hasFileTransfer(ev.dataTransfer)) return;
        const target = ev.currentTarget as HTMLElement;
        target.classList.remove("drag-over");
      });

      card.addEventListener("drop", (ev) => {
        if (hasFileTransfer(ev.dataTransfer)) return;
        if (isInteractiveElement(ev.target)) return;

        ev.preventDefault();
        const target = ev.currentTarget as HTMLElement;
        target.classList.remove("drag-over");

        let draggingId: string | null = null;

        try {
          draggingId = (ev.dataTransfer as DataTransfer).getData("text/plain");
        } catch {
          // Falls keine Drag-Payload lesbar ist, faellt der Code unten auf das aktive Drag-Element zurueck.
        }

        if (!draggingId) {
          const draggingEl = document.querySelector(".image-card.dragging") as HTMLElement | null;
          if (draggingEl) draggingId = draggingEl.getAttribute("data-id");
        }

        const targetId = target.getAttribute("data-id");

        if (draggingId && targetId && draggingId !== targetId) {
          const targetIndex = imageItems.findIndex((i) => i.id === targetId);
          if (targetIndex >= 0) {
            moveItemToPosition(draggingId, targetIndex + 1);
          }
        }
      });

      const position = document.createElement("div");
      position.className = "image-position";

      const numberBadge = document.createElement("div");
      numberBadge.className = "image-number";
      numberBadge.textContent = String(index + 1);
      position.appendChild(numberBadge);

      const controls = document.createElement("div");
      controls.className = "position-controls";

      const posInputLeft = document.createElement("input");
      posInputLeft.type = "number";
      posInputLeft.className = "pos-input";
      posInputLeft.min = "1";
      posInputLeft.value = String(item.position || index + 1);
      posInputLeft.draggable = false;
      posInputLeft.addEventListener("mousedown", (e) => e.stopPropagation());
      posInputLeft.addEventListener("click", (e) => e.stopPropagation());
      posInputLeft.addEventListener("change", () => {
        const newPos = Math.max(1, Math.floor(Number(posInputLeft.value) || 1));
        moveItemToPosition(item.id, newPos);
      });

      controls.appendChild(posInputLeft);
      position.appendChild(controls);

      const selectionWrap = document.createElement("div");
      selectionWrap.className = "selection-wrap";

      const selectLabel = document.createElement("label");
      selectLabel.className = "select-label";

      const selectCheckbox = document.createElement("input");
      selectCheckbox.type = "checkbox";
      selectCheckbox.checked = item.selected !== false;
      selectCheckbox.draggable = false;
      selectCheckbox.addEventListener("click", (e) => e.stopPropagation());
      selectCheckbox.addEventListener("change", () => {
        item.selected = selectCheckbox.checked;
        setMeta(item.hash, { selected: selectCheckbox.checked });
      });

      const selectLabelText = document.createElement("span");
      selectLabelText.textContent = "Ausgewählt";

      selectLabel.appendChild(selectCheckbox);
      selectLabel.appendChild(selectLabelText);
      selectionWrap.appendChild(selectLabel);
      position.appendChild(selectionWrap);

      const previewWrap = document.createElement("div");
      previewWrap.className = "preview-wrap";

      const preview = document.createElement("img");
      preview.className = "image-preview";
      preview.src = item.previewUrl;
      preview.alt = item.name;
      preview.draggable = false;

      previewWrap.appendChild(preview);

      const contentContainer = document.createElement("div");
      contentContainer.className = "image-content";

      const info = document.createElement("div");
      info.className = "image-info";

      const fileName = document.createElement("div");
      fileName.className = "file-name";
      fileName.textContent = item.name;

      const fileMeta = document.createElement("div");
      fileMeta.className = "file-meta";
      const fileType = item.name.split(".").pop() || "";
      fileMeta.textContent = `${formatFileSize(item.size)} · ${fileType}`;

      info.appendChild(fileName);
      info.appendChild(fileMeta);

      const exifRaw = item.exif?.dateTimeOriginal || "";
      let exifDate = "";
      let exifTime = "";

      if (exifRaw) {
        const parts = exifRaw.split(" ");
        exifDate = parts[0] ? parts[0].replace(/:/g, "-") : "";
        exifTime = parts[1] || "";
      }

      const model = item.exif?.model || "";

      const smallInfo = document.createElement("div");
      smallInfo.className = "small-info";

      if (exifDate) {
        const dateLine = document.createElement("div");
        dateLine.className = "info-line exif-date";
        dateLine.textContent = exifDate;
        smallInfo.appendChild(dateLine);
      }

      if (exifTime) {
        const timeLine = document.createElement("div");
        timeLine.className = "info-line exif-time";
        timeLine.textContent = exifTime;
        smallInfo.appendChild(timeLine);
      }

      if (model) {
        const modelLine = document.createElement("div");
        modelLine.className = "info-line camera-model";
        modelLine.textContent = model;
        smallInfo.appendChild(modelLine);
      }

      if (smallInfo.children.length > 0) {
        info.appendChild(smallInfo);
      }

      const captionContainer = document.createElement("div");
      captionContainer.className = "image-caption";

      const captionLabel = document.createElement("div");
      captionLabel.className = "caption-label";
      captionLabel.textContent = "Beschriftung";

      const captionInput = document.createElement("textarea");
      captionInput.className = "caption-input";
      captionInput.rows = 6;
      captionInput.placeholder = "Kurze Bildbeschriftung eingeben";
      captionInput.value = item.caption || "";
      captionInput.draggable = false;
      captionInput.addEventListener("mousedown", (e) => e.stopPropagation());
      captionInput.addEventListener("click", (e) => e.stopPropagation());

      captionInput.addEventListener("input", (event) => {
        const target = event.target as HTMLTextAreaElement;
        item.caption = target.value;

        setMeta(item.hash, {
          caption: target.value,
          position: item.position,
        });
      });

      captionContainer.appendChild(captionLabel);
      captionContainer.appendChild(captionInput);

      const actionsContainer = document.createElement("div");
      actionsContainer.className = "image-actions-right";

      const insertWithCaptionButton = buildActionButton({
        action: "insert-single-caption",
        id: item.id,
        variant: "primary",
        iconClass: "insert-button-icon-text",
        label: "Bild mit Beschriftung einfügen",
      });

      const insertButton = buildActionButton({
        action: "insert-single",
        id: item.id,
        variant: "quiet",
        iconClass: "insert-button-icon-image",
        label: "Bild nur mit Nummerierung einfügen",
      });
      const insertPlainButton = buildActionButton({
        action: "insert-single-plain",
        id: item.id,
        variant: "quiet",
        iconClass: "insert-button-icon-image-plain",
        label: "Bild ohne Beschriftung einfügen",
      });

      actionsContainer.appendChild(insertWithCaptionButton);
      actionsContainer.appendChild(insertButton);
      actionsContainer.appendChild(insertPlainButton);
      position.appendChild(actionsContainer);

      card.appendChild(position);
      card.appendChild(previewWrap);
      contentContainer.appendChild(info);
      contentContainer.appendChild(captionContainer);
      card.appendChild(contentContainer);
      card.appendChild(actionsContainer);

      imageList.appendChild(card);
    });
  }

  function moveItemToPosition(id: string, targetPos: number) {
    const idx = imageItems.findIndex((i) => i.id === id);
    if (idx === -1) return;

    const [item] = imageItems.splice(idx, 1);
    const insertIndex = Math.max(0, Math.min(targetPos - 1, imageItems.length));
    imageItems.splice(insertIndex, 0, item);

    if (sortMode !== "custom") {
      setCurrentSortMode("custom");
    }

    persistImagePositions();

    renderImageList();
  }

  function buildActionButton(options: {
    action: "insert-single" | "insert-single-caption" | "insert-single-plain";
    id: string;
    variant: "primary" | "quiet";
    iconClass:
      | "insert-button-icon-image"
      | "insert-button-icon-text"
      | "insert-button-icon-image-plain";
    label: string;
  }): HTMLButtonElement {
    const button = document.createElement("button");
    button.className = `insert-button icon-action-button ${
      options.variant === "primary" ? "insert-button-primary" : "insert-button-quiet"
    }`;
    button.type = "button";
    button.setAttribute("data-action", options.action);
    button.setAttribute("data-id", options.id);
    button.setAttribute("title", options.label);
    button.setAttribute("aria-label", options.label);
    button.draggable = false;

    const content = document.createElement("span");
    content.className = "insert-button-content";

    const icon = document.createElement("span");
    icon.className = `insert-button-icon ${options.iconClass}`;
    content.appendChild(icon);
    button.appendChild(content);

    return button;
  }

  imageList?.addEventListener("click", async (event) => {
    const target = event.target as HTMLElement | null;
    const button = target?.closest("button[data-action]") as HTMLButtonElement | null;

    if (!button) return;

    event.preventDefault();
    event.stopPropagation();

    const id = button.getAttribute("data-id");
    if (!id) {
      setStatus("Kein Bild gefunden.");
      return;
    }

    const item = imageItems.find((i) => i.id === id);
    if (!item) {
      setStatus("Bild nicht gefunden.");
      return;
    }

    try {
      if (typeof (window as any).Word === "undefined") {
        throw new Error("Word API is not available in this host (Word is undefined)");
      }

      // Stabiler Kernbereich: bestehendes Einzelfoto-Einfuegen unveraendert nutzen.
      const action = button.getAttribute("data-action");
      const captionMode: CaptionMode =
        action === "insert-single-caption"
          ? "full"
          : action === "insert-single"
            ? "numberOnly"
            : "plainImage";
      await insertSingleImageAtSelection(item, insertSizeCm, {
        includeCaption: captionMode !== "plainImage",
        captionMode,
      });
      setStatus(
        captionMode === "full"
          ? `Bild "${item.name}" wurde mit Beschriftung in Word eingefügt.`
          : captionMode === "numberOnly"
            ? `Bild "${item.name}" wurde nur mit Nummerierung in Word eingefügt.`
            : `Bild "${item.name}" wurde ohne Beschriftung in Word eingefügt.`
      );
    } catch (e: any) {
      console.error("Fehler beim Einfügen eines Bildes:", e);
      setStatus(`Fehler beim Word.run: ${e?.message || "Unbekannter Fehler"}`);
    }
  });

  pickFilesButton?.addEventListener("click", async () => {
    const nativePicker = (window as any).showOpenFilePicker;

    if (typeof nativePicker === "function") {
      try {
        const handles = await nativePicker({
          multiple: true,
          types: [
            {
              description: "Images",
              accept: {
                "image/jpeg": [".jpg", ".jpeg"],
                "image/png": [".png"],
                "image/webp": [".webp"],
                "image/tiff": [".tif", ".tiff"],
                "image/gif": [".gif"],
                "image/bmp": [".bmp"],
              },
            },
          ],
        });

        const files: File[] = [];

        for (const handle of handles) {
          try {
            const file = await handle.getFile();
            files.push(file);
          } catch (err) {
            console.warn("Fehler beim Lesen einer Datei vom Picker-Handle:", err);
          }
        }

        if (files.length > 0) {
          currentFolderHandle = null;
          await appendFiles(files);
        }

        return;
      } catch (err) {
        console.warn("showOpenFilePicker nicht verfügbar oder abgebrochen:", err);
      }
    }

    imageUpload?.click();
  });

  pickFolderButton?.addEventListener("click", async () => {
    const nativeDirPicker = (window as any).showDirectoryPicker;

    if (typeof nativeDirPicker === "function") {
      try {
        const dirHandle = await nativeDirPicker();
        currentFolderHandle = dirHandle;
        const files: File[] = [];

        async function traverseDirectory(handle: any, relativePrefix = "") {
          for await (const entry of handle.values()) {
            const entryPath = relativePrefix ? `${relativePrefix}/${entry.name}` : entry.name;

            if (entry.kind === "file") {
              try {
                const file: File = await entry.getFile();
                (file as File & { relativePath?: string }).relativePath = entryPath;
                files.push(file);
              } catch (err) {
                console.warn("Fehler beim Lesen einer Datei aus dem Ordner-Handle:", err);
              }
            } else if (entry.kind === "directory") {
              await traverseDirectory(entry, entryPath);
            }
          }
        }

        await traverseDirectory(dirHandle);

        if (files.length > 0) {
          await appendFiles(files);
        }

        const projectSummary = await autoLoadProjectFileFromCurrentFolder();
        if (projectSummary) {
          setStatus(projectSummary);
        }

        return;
      } catch (err) {
        console.warn("showDirectoryPicker nicht verfügbar oder abgebrochen:", err);
      }
    }

    folderUpload?.click();
  });

  dropZone?.addEventListener("dragenter", (event) => {
    if (!hasFileTransfer(event.dataTransfer)) {
      return;
    }

    event.preventDefault();
    dropZoneDragDepth += 1;
    setDropZoneActive(true);
  });

  dropZone?.addEventListener("dragover", (event) => {
    if (!hasFileTransfer(event.dataTransfer)) {
      return;
    }

    event.preventDefault();
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = "copy";
    }
    setDropZoneActive(true);
  });

  dropZone?.addEventListener("dragleave", (event) => {
    if (!hasFileTransfer(event.dataTransfer)) {
      return;
    }

    event.preventDefault();
    dropZoneDragDepth = Math.max(0, dropZoneDragDepth - 1);

    if (dropZoneDragDepth === 0) {
      setDropZoneActive(false);
    }
  });

  dropZone?.addEventListener("drop", async (event) => {
    if (!hasFileTransfer(event.dataTransfer)) {
      return;
    }

    event.preventDefault();
    resetDropZoneState();

    const droppedFiles = Array.from(event.dataTransfer?.files || []);
    if (droppedFiles.length === 0) {
      setStatus("Keine Dateien im Drop erkannt.");
      return;
    }

    currentFolderHandle = null;
    await appendFiles(droppedFiles);
  });

  document.addEventListener("dragover", (event) => {
    if (!hasFileTransfer(event.dataTransfer)) {
      return;
    }

    event.preventDefault();
  });

  document.addEventListener("drop", (event) => {
    if (!hasFileTransfer(event.dataTransfer)) {
      return;
    }

    event.preventDefault();
  });

  imageUpload?.addEventListener("change", async () => {
    if (!imageUpload.files || imageUpload.files.length === 0) {
      setStatus("Keine Dateien ausgewählt.");
      return;
    }

    currentFolderHandle = null;
    await appendFiles(imageUpload.files);
  });

  folderUpload?.addEventListener("change", async () => {
    if (!folderUpload.files || folderUpload.files.length === 0) {
      setStatus("Kein Ordner ausgewählt.");
      return;
    }

    currentFolderHandle = null;
    await appendFiles(Array.from(folderUpload.files));
  });

  saveProjectButton?.addEventListener("click", async () => {
    await saveProjectFile();
  });

  loadProjectButton?.addEventListener("click", async () => {
    const nativePicker = (window as any).showOpenFilePicker;

    if (typeof nativePicker === "function") {
      try {
        const handles = await nativePicker({
          multiple: false,
          types: [
            {
              description: "WERKHAUS Bilddaten",
              accept: {
                "application/json": [".json"],
              },
            },
          ],
        });

        const fileHandle = handles?.[0];
        if (!fileHandle) {
          setStatus("Keine Projektdatei ausgewählt.");
          return;
        }

        const file = await fileHandle.getFile();
        await loadProjectFromSelectedFile(file, "Projektdatei");
        return;
      } catch (error) {
        if ((error as { name?: string })?.name === "AbortError") {
          setStatus("Keine Projektdatei ausgewählt.");
          return;
        }

        console.warn("showOpenFilePicker nicht verfügbar oder abgebrochen:", error);
      }
    }

    projectImportUpload?.click();
  });

  projectImportUpload?.addEventListener("change", async () => {
    if (!projectImportUpload.files || projectImportUpload.files.length === 0) {
      setStatus("Keine Projektdatei ausgewählt.");
      return;
    }

    const file = projectImportUpload.files[0];
    projectImportUpload.value = "";
    await loadProjectFromSelectedFile(file, "Projektdatei");
  });

  clearImagesButton?.addEventListener("click", () => {
    clearImageList();
  });

  selectAllButton?.addEventListener("click", () => {
    setAllItemsSelected(true);
  });

  selectNoneButton?.addEventListener("click", () => {
    setAllItemsSelected(false);
  });

  insertSelectedButton?.addEventListener("click", async () => {
    await insertSelectedImages("numberOnly");
  });

  insertSelectedCaptionButton?.addEventListener("click", async () => {
    await insertSelectedImages("full");
  });

  insertSelectedPlainButton?.addEventListener("click", async () => {
    await insertSelectedImages("plainImage");
  });

  toggleInfoButton?.addEventListener("click", () => {
    showInfo = !showInfo;
    updateVisibilityUI();
  });

  toggleCaptionButton?.addEventListener("click", () => {
    showCaptions = !showCaptions;
    updateVisibilityUI();
  });

  expandAllSectionsButton?.addEventListener("click", () => {
    setAllSectionsCollapsed(false);
  });

  collapseAllSectionsButton?.addEventListener("click", () => {
    setAllSectionsCollapsed(true);
  });

  bindSectionToggleButtons();

  sortModeSelect?.addEventListener("change", () => {
    const nextMode = sortModeSelect.value as SortMode;
    if (nextMode === sortMode) {
      updateSortUI();
      return;
    }

    setCurrentSortMode(nextMode);
    applySortModeToItems();
  });

  updatePreviewSize();
  updateVisibilityUI();
  updateSortUI();
  restoreCollapsedSections();
  resetDropZoneState();
  persistImagePositions();
  renderImageList();
  setStatus("Bereit.");
}

if (
  typeof (window as any).Office !== "undefined" &&
  typeof (window as any).Office.onReady === "function"
) {
  (window as any).Office.onReady(initTaskpane);
} else {
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initTaskpane);
  } else {
    initTaskpane();
  }
}

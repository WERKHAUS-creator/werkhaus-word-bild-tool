import type { CaptionMode, ImageItem } from "./types";
import { IMAGE_FILE_INPUT_ACCEPT, importImageFiles } from "./importImages";
import { createTaskpanePersistence } from "./persistence";
import { insertSingleImageAtSelection } from "./wordInsert";

// Stabiler Kernbereich:
// Bildimport, Bildauswahl und Einzelfoto-Einfuegen bleiben hier bewusst getrennt.
// Nur bei nachgewiesenem Fehler an diesen Pfaden aendern.
let imageItems: ImageItem[] = [];
const DEFAULT_PREVIEW_SIZE_PX = 120;
const MIN_PREVIEW_SIZE_PX = 120;
const MAX_PREVIEW_SIZE_PX = 500;
const MAX_INSERT_SIZE_CM = 16;

function initTaskpane() {
  const { loadAllMeta, getMeta, setMeta } = createTaskpanePersistence();
  const statusElement = document.getElementById("statusMessage");
  const imageList = document.getElementById("imageList");
  const dropZone = document.getElementById("imageDropZone");

  const imageUpload = document.getElementById("imageUpload") as HTMLInputElement | null;
  const folderUpload = document.getElementById("folderUpload") as HTMLInputElement | null;

  const pickFilesButton = document.getElementById("pickFilesButton");
  const pickFolderButton = document.getElementById("pickFolderButton");
  const clearImagesButton = document.getElementById("clearImagesButton");
  const toggleInfoButton = document.getElementById("toggleInfoButton");
  const toggleCaptionButton = document.getElementById("toggleCaptionButton");
  const selectAllButton = document.getElementById("selectAllButton");
  const selectNoneButton = document.getElementById("selectNoneButton");
  const insertSelectedButton = document.getElementById("insertSelectedButton");
  const insertSelectedCaptionButton = document.getElementById("insertSelectedCaptionButton");
  const insertSelectedPlainButton = document.getElementById("insertSelectedPlainButton");

  const previewSizeRange = document.getElementById("previewSizeRange") as HTMLInputElement | null;
  const previewSizeValue = document.getElementById("previewSizeValue");
  const insertSizeRange = document.getElementById("insertSizeRange") as HTMLInputElement | null;
  const insertSizeValue = document.getElementById("insertSizeValue");

  loadAllMeta();

  let insertSizeCm = 10;
  let previewSizePx = DEFAULT_PREVIEW_SIZE_PX;

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
    return [...imageItems]
      .sort((a, b) => (a.position || 0) - (b.position || 0))
      .filter((item) => item.selected !== false);
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

      resetInputValue(imageUpload, "imageUpload");
      resetInputValue(folderUpload, "folderUpload");

      persistImagePositions();
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

    const itemsToRender = [...imageItems].sort((a, b) => (a.position || 0) - (b.position || 0));

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
        const files: File[] = [];

        async function traverseDirectory(handle: any) {
          for await (const entry of handle.values()) {
            if (entry.kind === "file") {
              try {
                const file: File = await entry.getFile();
                files.push(file);
              } catch (err) {
                console.warn("Fehler beim Lesen einer Datei aus dem Ordner-Handle:", err);
              }
            } else if (entry.kind === "directory") {
              await traverseDirectory(entry);
            }
          }
        }

        await traverseDirectory(dirHandle);

        if (files.length > 0) {
          await appendFiles(files);
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

    await appendFiles(imageUpload.files);
  });

  folderUpload?.addEventListener("change", async () => {
    if (!folderUpload.files || folderUpload.files.length === 0) {
      setStatus("Kein Ordner ausgewählt.");
      return;
    }

    await appendFiles(Array.from(folderUpload.files));
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

  updatePreviewSize();
  updateVisibilityUI();
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

import * as XLSX from "xlsx";

type VariableStatus = "not-inserted" | "up-to-date" | "outdated";
type DateFormatMode = "short" | "long";

type VariableItem = {
  name: string;
  ref: string;
  displayValue: string;
  matrix: string[][];
  isBlock: boolean;
  usageCount: number;
  status: VariableStatus;
};

type SavedFileMeta = {
  fileName: string;
  lastLoadedAt: string;
};

type ParsedControlMeta = {
  variableName: string;
  mode: "single" | "table";
  fitWidth: boolean;
  withWordBorders: boolean;
  keepFormattingOnUpdate: boolean;
  useExcelFormatting: boolean;
};

type TableOptions = {
  fitWidth: boolean;
  withWordBorders: boolean;
  keepFormattingOnUpdate: boolean;
  useExcelFormatting: boolean;
};

let loadedVariables: VariableItem[] = [];
let isHighlightActive = false;
let currentDateFormat: DateFormatMode = "short";
let currentExcelFile: File | null = null;
let tableOptionsByVariable: Record<string, TableOptions> = {};

Office.onReady(() => {
  const loadButton = document.getElementById("loadExcelButton") as HTMLButtonElement | null;
  const updateAllButton = document.getElementById("updateAllButton") as HTMLButtonElement | null;
  const toggleHighlightButton = document.getElementById("toggleHighlightButton") as HTMLButtonElement | null;
  const dateFormatSelect = document.getElementById("dateFormatSelect") as HTMLSelectElement | null;

  if (loadButton) loadButton.onclick = handleExcelLoad;
  if (updateAllButton) updateAllButton.onclick = handleUpdateAll;
  if (toggleHighlightButton) toggleHighlightButton.onclick = toggleHighlight;

  if (dateFormatSelect) {
    currentDateFormat = (dateFormatSelect.value as DateFormatMode) || "short";
    dateFormatSelect.onchange = async () => {
      currentDateFormat = (dateFormatSelect.value as DateFormatMode) || "short";
      if (currentExcelFile) {
        await loadExcelFile(currentExcelFile, false);
        setStatus("Datumsformat geändert und Excel-Datei neu geladen.");
      } else {
        setStatus("Datumsformat geändert.");
      }
    };
  }

  renderSavedFileInfo();
  renderHighlightButtonState();
});

async function handleExcelLoad(): Promise<void> {
  const fileInput = document.getElementById("excelFile") as HTMLInputElement | null;
  if (!fileInput) return;

  const file = fileInput.files?.[0];
  if (!file) {
    setStatus("Bitte zuerst eine Excel-Datei auswählen.");
    return;
  }

  currentExcelFile = file;
  await loadExcelFile(file, true);
}

async function loadExcelFile(file: File, saveMeta: boolean): Promise<void> {
  const variablesList = document.getElementById("variablesList") as HTMLDivElement | null;
  if (!variablesList) return;

  try {
    const arrayBuffer = await file.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
      cellNF: true,
      cellText: true,
      cellDates: true,
    });

    const result: VariableItem[] = [];
    const names = workbook.Workbook?.Names ?? [];

    for (const namedItem of names as any[]) {
      const variableName = String(namedItem.Name ?? "").trim();
      const ref = String(namedItem.Ref ?? "").trim();

      if (!variableName || !ref) continue;

      const resolved = resolveNamedReferenceData(workbook, ref);

      result.push({
        name: variableName,
        ref,
        displayValue: resolved.displayValue,
        matrix: resolved.matrix,
        isBlock: resolved.isBlock,
        usageCount: 0,
        status: "not-inserted",
      });

      if (resolved.isBlock && !tableOptionsByVariable[variableName]) {
        tableOptionsByVariable[variableName] = {
          fitWidth: true,
          withWordBorders: true,
          keepFormattingOnUpdate: true,
          useExcelFormatting: false,
        };
      }
    }

    loadedVariables = result;

    if (saveMeta) {
      await saveLinkedFileMeta({
        fileName: file.name,
        lastLoadedAt: new Date().toLocaleString("de-DE"),
      });
    }

    renderSavedFileInfo();

    if (result.length === 0) {
      variablesList.innerHTML = "<p>Keine benannten Variablen oder Bereiche gefunden.</p>";
      setStatus("Es wurden keine benannten Variablen gefunden.");
      return;
    }

    await refreshUsageStatus();
    setStatus(`${result.length} Variable(n) geladen.`);
  } catch (error) {
    console.error(error);
    setStatus("Fehler beim Lesen der Excel-Datei.");
  }
}

async function handleUpdateAll(): Promise<void> {
  if (loadedVariables.length === 0) {
    setStatus("Bitte zuerst eine Excel-Datei laden.");
    return;
  }

  try {
    let updatedCount = 0;

    await Word.run(async (context) => {
      const controls = context.document.contentControls;
      controls.load("items/tag");
      await context.sync();

      for (const control of controls.items) {
        const meta = parseControlMeta(control.tag);
        if (!meta) continue;

        const variable = loadedVariables.find((v) => v.name === meta.variableName);
        if (!variable) continue;

        await renderVariableIntoControl(context, control, variable, meta, true);
        updatedCount++;
      }

      await context.sync();
    });

    await refreshUsageStatus();

    if (isHighlightActive) {
      await applyHighlightToAllVariables(true);
    }

    setStatus(`${updatedCount} Feld(er) im Dokument aktualisiert.`);
  } catch (error) {
    console.error(error);
    setStatus("Fehler beim Aktualisieren der Variablen im Word-Dokument.");
  }
}

async function updateSingleVariable(variableName: string): Promise<void> {
  const variable = loadedVariables.find((v) => v.name === variableName);
  if (!variable) return;

  try {
    let updatedCount = 0;

    await Word.run(async (context) => {
      const controls = context.document.contentControls;
      controls.load("items/tag");
      await context.sync();

      for (const control of controls.items) {
        const meta = parseControlMeta(control.tag);
        if (!meta) continue;
        if (meta.variableName !== variableName) continue;

        await renderVariableIntoControl(context, control, variable, meta, true);
        updatedCount++;
      }

      await context.sync();
    });

    await refreshUsageStatus();

    if (isHighlightActive) {
      await applyHighlightToAllVariables(true);
    }

    setStatus(`${updatedCount} Vorkommen von "${variableName}" aktualisiert.`);
  } catch (error) {
    console.error(error);
    setStatus(`Fehler beim Aktualisieren von "${variableName}".`);
  }
}

async function refreshUsageStatus(): Promise<void> {
  if (loadedVariables.length === 0) {
    renderVariables(loadedVariables);
    return;
  }

  try {
    for (const variable of loadedVariables) {
      variable.usageCount = 0;
      variable.status = "not-inserted";
    }

    const variableMap = new Map<string, VariableItem>();
    loadedVariables.forEach((v) => variableMap.set(v.name, v));

    await Word.run(async (context) => {
      const controls = context.document.contentControls;
      controls.load("items/tag,text");
      await context.sync();

      for (const control of controls.items) {
        const meta = parseControlMeta(control.tag);
        if (!meta) continue;

        const variable = variableMap.get(meta.variableName);
        if (!variable) continue;

        variable.usageCount += 1;

        const expectedText =
          meta.mode === "single"
            ? variable.displayValue
            : variable.matrix.map((row) => row.join(" ")).join("\n");

        const currentText = control.text ?? "";

        const matches =
          meta.mode === "single"
            ? textsEqualForInlineStatus(currentText, expectedText)
            : textsEqualForBlockStatus(currentText, expectedText);

        if (variable.status !== "outdated") {
          variable.status = matches ? "up-to-date" : "outdated";
        }
      }

      for (const variable of loadedVariables) {
        if (variable.usageCount === 0) {
          variable.status = "not-inserted";
        }
      }

      await context.sync();
    });

    renderVariables(loadedVariables);
  } catch (error) {
    console.error(error);
    setStatus("Fehler beim Prüfen des Dokumentstatus.");
  }
}

async function toggleHighlight(): Promise<void> {
  try {
    isHighlightActive = !isHighlightActive;
    await applyHighlightToAllVariables(isHighlightActive);
    renderHighlightButtonState();

    setStatus(
      isHighlightActive
        ? "Alle Variablen im Dokument sind hervorgehoben."
        : "Die Hervorhebung der Variablen wurde entfernt."
    );
  } catch (error) {
    console.error(error);
    isHighlightActive = !isHighlightActive;
    renderHighlightButtonState();
    setStatus("Fehler beim Umschalten der Hervorhebung.");
  }
}

async function applyHighlightToAllVariables(enable: boolean): Promise<void> {
  await Word.run(async (context) => {
    const controls = context.document.contentControls;
    controls.load("items/tag");
    await context.sync();

    for (const control of controls.items) {
      const meta = parseControlMeta(control.tag);
      if (!meta) continue;

      const range = control.getRange();
      range.font.highlightColor = enable ? "#FFF59D" : null;
    }

    await context.sync();
  });
}

function renderHighlightButtonState(): void {
  const button = document.getElementById("toggleHighlightButton");
  const stateInfo = document.getElementById("highlightStateInfo");

  if (button) {
    button.textContent = isHighlightActive
      ? "Hervorhebung ausblenden"
      : "Variablen hervorheben";
  }

  if (stateInfo) {
    stateInfo.textContent = isHighlightActive
      ? "Hervorhebung: aktiv"
      : "Hervorhebung: aus";
  }
}

function renderVariables(variables: VariableItem[]): void {
  const variablesList = document.getElementById("variablesList") as HTMLDivElement | null;
  if (!variablesList) return;

  if (variables.length === 0) {
    variablesList.innerHTML = "Noch keine Variablen geladen.";
    return;
  }

  variablesList.innerHTML = variables
    .map((item, index) => {
      const statusLabel =
        item.status === "up-to-date"
          ? "Aktuell"
          : item.status === "outdated"
            ? "Veraltet"
            : "Neu";

      const statusColor =
        item.status === "up-to-date"
          ? "#107c10"
          : item.status === "outdated"
            ? "#d83b01"
            : "#666";

      const preview = item.isBlock
        ? escapeHtml(item.displayValue).replace(/\n/g, "<br/>")
        : escapeHtml(item.displayValue);

      const savedOptions = tableOptionsByVariable[item.name] || {
        fitWidth: true,
        withWordBorders: true,
        keepFormattingOnUpdate: true,
        useExcelFormatting: false,
      };

      const optionsHtml = item.isBlock
        ? `
          <div style="margin-top:10px; display:grid; gap:8px;">
            <label style="font-size:12px; color:#444;">
              <input type="checkbox" id="excelfmt-${index}" ${savedOptions.useExcelFormatting ? "checked" : ""} />
              Excel-Format übernehmen
            </label>
            <label style="font-size:12px; color:#444;">
              <input type="checkbox" id="fit-${index}" ${savedOptions.fitWidth ? "checked" : ""} />
              Seitenbreite anpassen
            </label>
            <label style="font-size:12px; color:#444;">
              <input type="checkbox" id="borders-${index}" ${savedOptions.withWordBorders ? "checked" : ""} />
              Word-Rahmen erstellen
            </label>
            <label style="font-size:12px; color:#444;">
              <input type="checkbox" id="keepfmt-${index}" ${savedOptions.keepFormattingOnUpdate ? "checked" : ""} />
              Formatierung bei Aktualisierung beibehalten
            </label>
          </div>
        `
        : ``;

      return `
        <div style="margin-bottom:12px; padding:10px; border:1px solid #ccc; border-radius:4px;">
          <div style="font-size:13px; line-height:1.4;">
            <strong>${escapeHtml(item.name)}</strong>
            &nbsp;·&nbsp;Verwendet: ${item.usageCount}
            &nbsp;·&nbsp;${escapeHtml(item.ref)}
            &nbsp;·&nbsp;<span style="color:${statusColor}; font-weight:600;">${statusLabel}</span>
          </div>

          <div style="margin-top:8px; padding:8px; background:#f7f7f7; border-radius:4px; white-space:pre-wrap;">${preview}</div>

          ${optionsHtml}

          <div style="display:flex; gap:8px; margin-top:10px; flex-wrap:wrap;">
            <button data-index="${index}" class="insertVariableButton ms-Button ms-Button--primary">
              <span class="ms-Button-label">Einfügen</span>
            </button>

            <button data-index="${index}" class="updateSingleVariableButton ms-Button">
              <span class="ms-Button-label">Nur diese aktualisieren</span>
            </button>
          </div>
        </div>
      `;
    })
    .join("");

  const insertButtons = document.querySelectorAll(".insertVariableButton");
  insertButtons.forEach((button) => {
    button.addEventListener("click", async (event) => {
      const target = event.currentTarget as HTMLButtonElement;
      const index = Number(target.getAttribute("data-index"));
      const variable = loadedVariables[index];
      if (!variable) return;

      const meta = getMetaFromUi(index, variable);
      await insertVariable(variable, meta);
    });
  });

  const updateButtons = document.querySelectorAll(".updateSingleVariableButton");
  updateButtons.forEach((button) => {
    button.addEventListener("click", async (event) => {
      const target = event.currentTarget as HTMLButtonElement;
      const index = Number(target.getAttribute("data-index"));
      const variable = loadedVariables[index];
      if (!variable) return;

      const meta = getMetaFromUi(index, variable);
      tableOptionsByVariable[variable.name] = {
        fitWidth: meta.fitWidth,
        withWordBorders: meta.withWordBorders,
        keepFormattingOnUpdate: meta.keepFormattingOnUpdate,
        useExcelFormatting: meta.useExcelFormatting,
      };

      await updateSingleVariable(variable.name);
    });
  });
}

function getMetaFromUi(index: number, variable: VariableItem): ParsedControlMeta {
  if (!variable.isBlock) {
    return {
      variableName: variable.name,
      mode: "single",
      fitWidth: false,
      withWordBorders: false,
      keepFormattingOnUpdate: true,
      useExcelFormatting: false,
    };
  }

  const excelFmtCheckbox = document.getElementById(`excelfmt-${index}`) as HTMLInputElement | null;
  const fitCheckbox = document.getElementById(`fit-${index}`) as HTMLInputElement | null;
  const bordersCheckbox = document.getElementById(`borders-${index}`) as HTMLInputElement | null;
  const keepFmtCheckbox = document.getElementById(`keepfmt-${index}`) as HTMLInputElement | null;

  const meta: ParsedControlMeta = {
    variableName: variable.name,
    mode: "table",
    fitWidth: !!fitCheckbox?.checked,
    withWordBorders: !!bordersCheckbox?.checked,
    keepFormattingOnUpdate: !!keepFmtCheckbox?.checked,
    useExcelFormatting: !!excelFmtCheckbox?.checked,
  };

  tableOptionsByVariable[variable.name] = {
    fitWidth: meta.fitWidth,
    withWordBorders: meta.withWordBorders,
    keepFormattingOnUpdate: meta.keepFormattingOnUpdate,
    useExcelFormatting: meta.useExcelFormatting,
  };

  return meta;
}

async function insertVariable(variable: VariableItem, meta: ParsedControlMeta): Promise<void> {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const control = selection.insertContentControl();
      control.tag = buildTag(meta);
      control.title = `Excel-Variable: ${variable.name}`;
      control.appearance = "BoundingBox";
      control.cannotDelete = false;
      control.cannotEdit = false;
      control.placeholderText = "";

      await renderVariableIntoControl(context, control, variable, meta, false);

      if (meta.mode === "single") {
        await moveCursorOutsideSingleControl(context, control);
      } else {
        const cursorRange = control.getRange("After");
        cursorRange.select();
      }

      await context.sync();
    });

    await refreshUsageStatus();

    if (isHighlightActive) {
      await applyHighlightToAllVariables(true);
    }

    setStatus(`Variable "${variable.name}" wurde eingefügt.`);
  } catch (error) {
    console.error(error);
    setStatus(`Fehler beim Einfügen der Variable "${variable.name}".`);
  }
}

async function moveCursorOutsideSingleControl(context: Word.RequestContext, control: Word.ContentControl): Promise<void> {
  const afterRange = control.getRange("After");
  afterRange.load("text");
  await context.sync();

  const nextText = afterRange.text ?? "";

  if (nextText.length === 0) {
    afterRange.insertText(" ", Word.InsertLocation.start);
    await context.sync();
  } else {
    const firstChar = nextText.charAt(0);
    if (!/\s/.test(firstChar) && !/^[,.;:!?)}\]»]/.test(firstChar)) {
      afterRange.insertText(" ", Word.InsertLocation.start);
      await context.sync();
    }
  }

  const finalAfter = control.getRange("After");
  finalAfter.select();
}

async function renderVariableIntoControl(
  context: Word.RequestContext,
  control: Word.ContentControl,
  variable: VariableItem,
  meta: ParsedControlMeta,
  isUpdate: boolean
): Promise<void> {
  if (meta.mode === "table" && variable.isBlock) {
    if (isUpdate && meta.keepFormattingOnUpdate) {
      const range = control.getRange();
      const tables = range.tables;
      tables.load("items");
      await context.sync();

      if (tables.items.length > 0) {
        const table = tables.items[0];
        const existingRowCount = table.rowCount;
        const existingColCount = table.getCell(0, 0) ? variable.matrix[0]?.length ?? 0 : variable.matrix[0]?.length ?? 0;

        if (existingRowCount === variable.matrix.length) {
          for (let r = 0; r < variable.matrix.length; r++) {
            const row = variable.matrix[r];
            for (let c = 0; c < row.length; c++) {
              const cell = table.getCell(r, c);
              cell.body.insertText(row[c], Word.InsertLocation.replace);
            }
          }
          return;
        }
      }
    }

    control.clear();
    const html = buildHtmlTable(
      variable.matrix,
      meta.fitWidth,
      meta.withWordBorders,
      meta.useExcelFormatting
    );
    control.insertHtml(html, Word.InsertLocation.replace);
    return;
  }

  control.insertText(variable.displayValue, Word.InsertLocation.replace);
}

function buildHtmlTable(
  matrix: string[][],
  fitWidth: boolean,
  withWordBorders: boolean,
  useExcelFormatting: boolean
): string {
  const tableWidth = fitWidth ? "width:100%;" : "width:auto;";
  const borderStyle = withWordBorders ? "1px solid #666" : "none";
  const cellPadding = "6px 8px";
  const extraStyle = useExcelFormatting ? "font-family:Calibri, Arial, sans-serif; font-size:11pt;" : "";

  const rowsHtml = matrix
    .map(
      (row) =>
        "<tr>" +
        row
          .map(
            (cell) =>
              `<td style="border:${borderStyle}; padding:${cellPadding}; vertical-align:top; ${extraStyle}">${escapeHtml(
                cell
              )}</td>`
          )
          .join("") +
        "</tr>"
    )
    .join("");

  return `<table style="border-collapse:collapse;${tableWidth}${extraStyle}">${rowsHtml}</table>`;
}

function parseControlMeta(tag: string | undefined): ParsedControlMeta | null {
  if (!tag) return null;
  if (!tag.startsWith("excelvar|")) return null;

  const parts = tag.split("|");
  if (parts.length < 7) return null;

  return {
    variableName: decodeSafe(parts[1]),
    mode: decodeSafe(parts[2]) === "table" ? "table" : "single",
    fitWidth: decodeSafe(parts[3]) === "1",
    withWordBorders: decodeSafe(parts[4]) === "1",
    keepFormattingOnUpdate: decodeSafe(parts[5]) === "1",
    useExcelFormatting: decodeSafe(parts[6]) === "1",
  };
}

function buildTag(meta: ParsedControlMeta): string {
  return [
    "excelvar",
    encodeURIComponent(meta.variableName),
    encodeURIComponent(meta.mode),
    encodeURIComponent(meta.fitWidth ? "1" : "0"),
    encodeURIComponent(meta.withWordBorders ? "1" : "0"),
    encodeURIComponent(meta.keepFormattingOnUpdate ? "1" : "0"),
    encodeURIComponent(meta.useExcelFormatting ? "1" : "0"),
  ].join("|");
}

function decodeSafe(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function textsEqualForInlineStatus(currentText: string, expectedText: string): boolean {
  return normalizeInlineText(currentText) === normalizeInlineText(expectedText);
}

function textsEqualForBlockStatus(currentText: string, expectedText: string): boolean {
  return normalizeBlockText(currentText) === normalizeBlockText(expectedText);
}

function normalizeInlineText(value: string): string {
  return value.replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeBlockText(value: string): string {
  return value
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\u00A0/g, " ")
    .replace(/\t/g, " ")
    .split("\n")
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter((line) => line.length > 0)
    .join("\n");
}

function resolveNamedReferenceData(
  workbook: XLSX.WorkBook,
  ref: string
): { displayValue: string; matrix: string[][]; isBlock: boolean } {
  const cleanedRef = ref.replace(/^=/, "");
  const match = cleanedRef.match(/^(?:'([^']+)'|([^!]+))!(.+)$/);

  if (!match) {
    return {
      displayValue: ref,
      matrix: [[ref]],
      isBlock: false,
    };
  }

  const sheetName = (match[1] || match[2] || "").trim();
  const rangeAddress = match[3];
  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    return {
      displayValue: ref,
      matrix: [[ref]],
      isBlock: false,
    };
  }

  try {
    const range = XLSX.utils.decode_range(rangeAddress);
    const matrix: string[][] = [];

    for (let r = range.s.r; r <= range.e.r; r++) {
      const row: string[] = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const cell = worksheet[cellAddress];
        row.push(getCellDisplayValue(cell));
      }
      matrix.push(row);
    }

    const isBlock = matrix.length > 1 || (matrix[0] && matrix[0].length > 1);
    const displayValue = isBlock
      ? matrix.map((row) => row.join(" | ")).join("\n")
      : (matrix[0]?.[0] ?? "");

    return { displayValue, matrix, isBlock };
  } catch {
    return {
      displayValue: ref,
      matrix: [[ref]],
      isBlock: false,
    };
  }
}

function getCellDisplayValue(cell: XLSX.CellObject | undefined): string {
  if (!cell) return "";

  const anyCell = cell as any;

  if (anyCell.t === "d" && anyCell.v instanceof Date) {
    return formatGermanDate(anyCell.v);
  }

  if (typeof anyCell.v === "number" && typeof anyCell.z === "string") {
    const formatString = anyCell.z.toLowerCase();

    if (formatString.includes("%")) {
      return new Intl.NumberFormat("de-DE", {
        style: "percent",
        minimumFractionDigits: guessFractionDigitsFromFormat(formatString),
        maximumFractionDigits: guessFractionDigitsFromFormat(formatString),
      }).format(anyCell.v);
    }

    if (looksLikeDateFormat(formatString)) {
      const parsedDate = XLSX.SSF.parse_date_code(anyCell.v);
      if (parsedDate) {
        const jsDate = new Date(parsedDate.y, parsedDate.m - 1, parsedDate.d);
        return formatGermanDate(jsDate);
      }
    }
  }

  if (typeof anyCell.w === "string" && anyCell.w.length > 0) {
    const maybeDate = tryNormalizeExcelDateToGerman(anyCell.w);
    if (maybeDate) return maybeDate;

    const maybePercent = tryNormalizePercentToGerman(anyCell.w);
    if (maybePercent) return maybePercent;

    return String(anyCell.w);
  }

  if (anyCell.v === undefined || anyCell.v === null) return "";
  return String(anyCell.v);
}

function looksLikeDateFormat(formatString: string): boolean {
  return /[dmy]/i.test(formatString);
}

function guessFractionDigitsFromFormat(formatString: string): number {
  const match = formatString.match(/0\.([0#]+)/);
  return match ? match[1].length : 0;
}

function formatGermanDate(date: Date): string {
  if (currentDateFormat === "long") {
    const day = String(date.getDate()).padStart(2, "0");
    const monthName = new Intl.DateTimeFormat("de-DE", { month: "long" }).format(date);
    const year = String(date.getFullYear());
    return `${day}. ${monthName} ${year}`;
  }

  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = String(date.getFullYear());
  return `${day}.${month}.${year}`;
}

function tryNormalizeExcelDateToGerman(value: string): string | null {
  const trimmed = value.trim();

  const isoMatch = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]) - 1;
    const day = Number(isoMatch[3]);
    return formatGermanDate(new Date(year, month, day));
  }

  const usMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (usMatch) {
    const month = Number(usMatch[1]) - 1;
    const day = Number(usMatch[2]);
    const year = Number(usMatch[3]);
    return formatGermanDate(new Date(year, month, day));
  }

  const deMatch = trimmed.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (deMatch) {
    const day = Number(deMatch[1]);
    const month = Number(deMatch[2]) - 1;
    let year = Number(deMatch[3]);
    if (year < 100) year += 2000;
    return formatGermanDate(new Date(year, month, day));
  }

  return null;
}

function tryNormalizePercentToGerman(value: string): string | null {
  const trimmed = value.trim();
  const match = trimmed.match(/^(-?\d+(?:[.,]\d+)?)\s?%$/);
  if (!match) return null;

  const num = Number(match[1].replace(",", "."));
  if (Number.isNaN(num)) return null;

  return `${num.toLocaleString("de-DE")} %`;
}

async function saveLinkedFileMeta(meta: SavedFileMeta): Promise<void> {
  Office.context.document.settings.set("excelWordLinkedFileMeta", meta);

  await new Promise<void>((resolve, reject) => {
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error);
      }
    });
  });
}

function renderSavedFileInfo(): void {
  const linkedFileInfo = document.getElementById("linkedFileInfo");
  if (!linkedFileInfo) return;

  const meta = Office.context.document.settings.get("excelWordLinkedFileMeta") as SavedFileMeta | undefined;

  if (!meta) {
    linkedFileInfo.innerHTML = "Noch keine Excel-Datei diesem Word-Dokument zugeordnet.";
    return;
  }

  linkedFileInfo.innerHTML =
    `<strong>Zugeordnete Excel-Datei:</strong> ${escapeHtml(meta.fileName)}<br/>` +
    `<strong>Letzter Import:</strong> ${escapeHtml(meta.lastLoadedAt)}`;
}

function setStatus(message: string): void {
  const status = document.getElementById("statusMessage");
  if (status) status.textContent = message;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
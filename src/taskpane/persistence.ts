import type { SortMode, StoredMeta } from "./types";

const META_KEY = "werkhaus-image-metadata";
const SETTINGS_KEY = "werkhaus-taskpane-settings";

// Stabiler Kernbereich:
// Lokale Metadaten fuer Caption, Reihenfolge und Caption-Aktivierung.
// Keine funktionale Erweiterung hier ohne Ruecksprache.

export interface TaskpanePersistence {
  loadAllMeta(): Record<string, StoredMeta>;
  loadSettings(): Record<string, unknown>;
  getMeta(hash: string): StoredMeta;
  getMetaKeys(): string[];
  getCollapsedSections(): string[];
  getSortMode(): SortMode | undefined;
  setMeta(hash: string, meta: StoredMeta): void;
  setCollapsedSections(sectionKeys: string[]): void;
  setSortMode(sortMode: SortMode): void;
}

export function createTaskpanePersistence(): TaskpanePersistence {
  let imageMetadataStore: Record<string, StoredMeta> = {};
  let taskpaneSettingsStore: { sortMode?: SortMode; collapsedSections?: string[] } = {};
  let metaSaveTimeout: number | null = null;
  let settingsSaveTimeout: number | null = null;

  function loadAllMeta(): Record<string, StoredMeta> {
    try {
      const raw = localStorage.getItem(META_KEY);
      if (!raw) {
        imageMetadataStore = {};
        return imageMetadataStore;
      }

      imageMetadataStore = JSON.parse(raw);
      return imageMetadataStore;
    } catch (e) {
      console.warn("Fehler beim Laden der Metadaten:", e);
      imageMetadataStore = {};
      return imageMetadataStore;
    }
  }

  function loadSettings(): Record<string, unknown> {
    try {
      const raw = localStorage.getItem(SETTINGS_KEY);
      if (!raw) {
        taskpaneSettingsStore = {};
        return taskpaneSettingsStore;
      }

      taskpaneSettingsStore = JSON.parse(raw);
      return taskpaneSettingsStore;
    } catch (e) {
      console.warn("Fehler beim Laden der Einstellungen:", e);
      taskpaneSettingsStore = {};
      return taskpaneSettingsStore;
    }
  }

  function saveAllMeta() {
    try {
      localStorage.setItem(META_KEY, JSON.stringify(imageMetadataStore));
    } catch (e) {
      console.warn("Fehler beim Speichern der Metadaten:", e);
    }
  }

  function saveAllSettings() {
    try {
      localStorage.setItem(SETTINGS_KEY, JSON.stringify(taskpaneSettingsStore));
    } catch (e) {
      console.warn("Fehler beim Speichern der Einstellungen:", e);
    }
  }

  function getMeta(hash: string): StoredMeta {
    return imageMetadataStore[hash] || {};
  }

  function getSortMode(): SortMode | undefined {
    return taskpaneSettingsStore.sortMode;
  }

  function getCollapsedSections(): string[] {
    return Array.isArray(taskpaneSettingsStore.collapsedSections)
      ? taskpaneSettingsStore.collapsedSections.filter((sectionKey): sectionKey is string => {
          return typeof sectionKey === "string" && sectionKey.length > 0;
        })
      : [];
  }

  function getMetaKeys(): string[] {
    return Object.keys(imageMetadataStore);
  }

  function scheduleSaveMeta() {
    if (metaSaveTimeout) {
      clearTimeout(metaSaveTimeout);
    }

    metaSaveTimeout = window.setTimeout(() => {
      try {
        saveAllMeta();
      } catch (e) {
        console.warn("Fehler beim geplanten Speichern der Metadaten:", e);
      }
      metaSaveTimeout = null;
    }, 200);
  }

  function setMeta(hash: string, meta: StoredMeta) {
    imageMetadataStore[hash] = { ...(imageMetadataStore[hash] || {}), ...meta };
    scheduleSaveMeta();
  }

  function scheduleSaveSettings() {
    if (settingsSaveTimeout) {
      clearTimeout(settingsSaveTimeout);
    }

    settingsSaveTimeout = window.setTimeout(() => {
      try {
        saveAllSettings();
      } catch (e) {
        console.warn("Fehler beim geplanten Speichern der Einstellungen:", e);
      }
      settingsSaveTimeout = null;
    }, 200);
  }

  function setSortMode(sortMode: SortMode) {
    taskpaneSettingsStore.sortMode = sortMode;
    scheduleSaveSettings();
  }

  function setCollapsedSections(sectionKeys: string[]) {
    taskpaneSettingsStore.collapsedSections = sectionKeys.filter(
      (sectionKey) => sectionKey.length > 0
    );
    scheduleSaveSettings();
  }

  return {
    loadAllMeta,
    loadSettings,
    getMeta,
    getMetaKeys,
    getCollapsedSections,
    getSortMode,
    setMeta,
    setCollapsedSections,
    setSortMode,
  };
}

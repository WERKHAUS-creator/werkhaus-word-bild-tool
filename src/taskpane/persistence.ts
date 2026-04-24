import type { StoredMeta } from "./types";

const META_KEY = "werkhaus-image-metadata";

// Stabiler Kernbereich:
// Lokale Metadaten fuer Caption, Reihenfolge und Caption-Aktivierung.
// Keine funktionale Erweiterung hier ohne Ruecksprache.

export interface TaskpanePersistence {
  loadAllMeta(): Record<string, StoredMeta>;
  getMeta(hash: string): StoredMeta;
  setMeta(hash: string, meta: StoredMeta): void;
}

export function createTaskpanePersistence(): TaskpanePersistence {
  let imageMetadataStore: Record<string, StoredMeta> = {};
  let metaSaveTimeout: number | null = null;

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

  function saveAllMeta() {
    try {
      localStorage.setItem(META_KEY, JSON.stringify(imageMetadataStore));
    } catch (e) {
      console.warn("Fehler beim Speichern der Metadaten:", e);
    }
  }

  function getMeta(hash: string): StoredMeta {
    return imageMetadataStore[hash] || {};
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

  return {
    loadAllMeta,
    getMeta,
    setMeta,
  };
}

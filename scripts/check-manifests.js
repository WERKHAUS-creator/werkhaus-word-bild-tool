const fs = require("fs");

const LOCAL_MANIFEST = "manifest.local.xml";
const PRODUCTION_MANIFEST = "manifest.production.xml";
const LOCAL_URL = "https://localhost:3001";
const PRODUCTION_URL = "https://tool2.wh-sv.de";

function readManifest(path) {
  if (!fs.existsSync(path)) {
    throw new Error(`${path} fehlt.`);
  }

  return fs.readFileSync(path, "utf8");
}

function assertContains(content, expected, path) {
  if (!content.includes(expected)) {
    throw new Error(`${path} enthaelt nicht ${expected}.`);
  }
}

function assertDoesNotContain(content, forbidden, path) {
  if (content.includes(forbidden)) {
    throw new Error(`${path} enthaelt unerwartet ${forbidden}.`);
  }
}

const localManifest = readManifest(LOCAL_MANIFEST);
const productionManifest = readManifest(PRODUCTION_MANIFEST);

assertContains(localManifest, LOCAL_URL, LOCAL_MANIFEST);
assertDoesNotContain(localManifest, PRODUCTION_URL, LOCAL_MANIFEST);

assertContains(productionManifest, PRODUCTION_URL, PRODUCTION_MANIFEST);
assertDoesNotContain(productionManifest, LOCAL_URL, PRODUCTION_MANIFEST);

console.log("Manifest-Struktur ok.");

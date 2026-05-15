const fs = require("fs");

const TEMPLATE_MANIFEST = "manifest.template.xml";
const DEV_MANIFEST = "manifest.dev.xml";
const PRODUCTION_MANIFEST = "manifest.xml";
const DEV_URL = "https://localhost:3001";
const PRODUCTION_URL = "https://tool2.wh-sv.de";
const PLACEHOLDER_URL = "__BASE_URL__";
const PLACEHOLDER_ID = "__APP_ID__";

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

const localManifest = readManifest(DEV_MANIFEST);
const productionManifest = readManifest(PRODUCTION_MANIFEST);
const templateManifest = readManifest(TEMPLATE_MANIFEST);

assertContains(templateManifest, PLACEHOLDER_URL, TEMPLATE_MANIFEST);
assertContains(templateManifest, PLACEHOLDER_ID, TEMPLATE_MANIFEST);
assertDoesNotContain(templateManifest, DEV_URL, TEMPLATE_MANIFEST);
assertDoesNotContain(templateManifest, PRODUCTION_URL, TEMPLATE_MANIFEST);

assertContains(localManifest, DEV_URL, DEV_MANIFEST);
assertDoesNotContain(localManifest, PRODUCTION_URL, DEV_MANIFEST);
assertDoesNotContain(localManifest, PLACEHOLDER_URL, DEV_MANIFEST);
assertDoesNotContain(localManifest, PLACEHOLDER_ID, DEV_MANIFEST);

assertContains(productionManifest, PRODUCTION_URL, PRODUCTION_MANIFEST);
assertDoesNotContain(productionManifest, DEV_URL, PRODUCTION_MANIFEST);
assertDoesNotContain(productionManifest, PLACEHOLDER_URL, PRODUCTION_MANIFEST);
assertDoesNotContain(productionManifest, PLACEHOLDER_ID, PRODUCTION_MANIFEST);

if (localManifest === productionManifest) {
  throw new Error("DEV- und Produktionsmanifest duerfen nicht identisch sein.");
}

const devVersionMatch = localManifest.match(/<Version>([^<]+)<\/Version>/);
const prodVersionMatch = productionManifest.match(/<Version>([^<]+)<\/Version>/);

if (!devVersionMatch || !prodVersionMatch) {
  throw new Error("Manifest-Versionen konnten nicht gelesen werden.");
}

if (devVersionMatch[1] !== prodVersionMatch[1]) {
  throw new Error("DEV- und Produktionsmanifest muessen dieselbe Version haben.");
}

const devIdMatch = localManifest.match(/<Id>([^<]+)<\/Id>/);
const prodIdMatch = productionManifest.match(/<Id>([^<]+)<\/Id>/);

if (!devIdMatch || !prodIdMatch) {
  throw new Error("Manifest-IDs konnten nicht gelesen werden.");
}

if (devIdMatch[1] === prodIdMatch[1]) {
  throw new Error("DEV- und Produktionsmanifest muessen unterschiedliche App-IDs haben.");
}

console.log("Manifest-Struktur ok.");

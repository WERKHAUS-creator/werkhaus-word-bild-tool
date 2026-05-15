const fs = require("fs");
const path = require("path");

const TEMPLATE_PATH = path.join(__dirname, "..", "manifest.template.xml");
const OUTPUTS = [
  {
    file: "manifest.dev.xml",
    appId: "6e810989-dd0a-4217-b841-503d28b8a05a",
    baseUrl: "https://localhost:3001",
  },
  {
    file: "manifest.xml",
    appId: "c2c1d8a4-4d33-4e0d-b4c2-7b6d2f9a8e41",
    baseUrl: "https://tool2.wh-sv.de",
  },
];

const MANIFEST_VERSION = "1.0.0.4";

function renderManifest(template, values) {
  return template
    .replace(/__APP_ID__/g, values.appId)
    .replace(/__MANIFEST_VERSION__/g, MANIFEST_VERSION)
    .replace(/__BASE_URL__/g, values.baseUrl);
}

function main() {
  const template = fs.readFileSync(TEMPLATE_PATH, "utf8");

  for (const output of OUTPUTS) {
    const content = renderManifest(template, output);
    const targetPath = path.join(__dirname, "..", output.file);
    fs.writeFileSync(targetPath, content, "utf8");
  }
}

main();

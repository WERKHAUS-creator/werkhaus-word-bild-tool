# Manifest Files

## Source of truth
- `manifest.template.xml` is the only central manifest source.
- `scripts/generate-manifest.js` generates the working manifests from that template.

## Generated manifests
- `manifest.dev.xml` is generated for local Word testing and points to `https://localhost:3001`.
- `manifest.xml` is generated for production publishing and points to `https://tool2.wh-sv.de`.
- `npm run start`, `npm run build`, `npm run validate:dev`, and `npm run validate:prod` generate the manifests automatically.
- `npm run build` also copies `manifest.xml` to `dist/manifest.xml`.

## Checks
- `npm run check:manifests` verifies that the template and generated manifests stay in sync.
- `npm run validate:dev` validates the local manifest with the Office validator.
- `npm run validate:prod` validates the production manifest with the Office validator.

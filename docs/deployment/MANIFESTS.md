# Manifest Files

## Source of truth
- `manifest.xml` is the production manifest source.
- `manifest.dev.xml` is the local development manifest source.

## Generated manifests
- `manifest.dev.xml` points to `https://localhost:3001`.
- `manifest.xml` points to `https://tool2.wh-sv.de`.
- `SupportUrl` points to `support.html` on the same base URL.
- `npm run start`, `npm run build`, `npm run validate:dev`, and `npm run validate:prod` use the manifests directly.
- `npm run build` also copies `manifest.xml` to `dist/manifest.xml`.

## Checks
- `npm run check:manifests` verifies that the local and production manifests stay in sync.
- `npm run validate:dev` validates the local manifest with the Office validator.
- `npm run validate:prod` validates the production manifest with the Office validator.

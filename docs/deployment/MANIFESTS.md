# Manifest Files

## Local testing
- Use `manifest.local.xml` for local Word testing.
- `npm run start` and `npm run stop` use `manifest.local.xml`.
- This manifest must point to `https://localhost:3001`.

## Production publishing
- Use `manifest.production.xml` for GitHub, Azure, and Microsoft 365 Admin Center publishing.
- This manifest must point to `https://tool2.wh-sv.de`.
- `npm run build` copies `manifest.production.xml` to `dist/manifest.xml`.
- The old root `manifest.xml` was intentionally removed to avoid confusion. `dist/manifest.xml` exists only as a build output.

## Checks
- `npm run check:manifests` verifies that local and production URLs are not mixed.
- `npm run validate:local` validates the local manifest with the Office validator.
- `npm run validate:prod` validates the production manifest with the Office validator.

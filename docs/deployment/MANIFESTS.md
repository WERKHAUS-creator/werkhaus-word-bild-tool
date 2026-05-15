# Manifest Files

## Local testing
- Use `manifest.dev.xml` for local Word testing.
- `npm run start` and `npm run stop` use `manifest.dev.xml`.
- This manifest must point to `https://localhost:3001`.

## Production publishing
- Use `manifest.xml` for GitHub, Azure, and Microsoft 365 Admin Center publishing.
- This manifest must point to `https://tool2.wh-sv.de`.
- `npm run build` copies `manifest.xml` to `dist/manifest.xml`.
- `dist/manifest.xml` exists only as a build output.

## Checks
- `npm run check:manifests` verifies that local and production URLs are not mixed.
- `npm run validate:dev` validates the local manifest with the Office validator.
- `npm run validate:prod` validates the production manifest with the Office validator.

# MS 365 Installation Readiness

This project currently contains a local development manifest (`manifests/manifest.dev.xml`) pointing to `https://localhost:3001`.

## Intended split
- Local/dev manifest: keep localhost URLs for development and sideloading.
- Production manifest: use hosted Azure URLs for enterprise rollout.

## Required production preparation before tenant-wide installation
1. Host stable build on Azure Static Web Apps from `main`.
2. Maintain the production manifest file `manifests/manifest.prod.xml` with Azure URLs.
3. Validate production manifest with `npm run validate` (or Office manifest validator).
4. Admin-center deployment should use production manifest only.

## Do not
- Do not deploy `manifests/manifest.dev.xml` (localhost) to tenant-wide production.
- Do not mix test URLs and production URLs in the same release manifest.

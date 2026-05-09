# MS 365 Installation Readiness

This project contains a local development manifest (`manifest.local.xml`) pointing to `https://localhost:3001` and a production manifest (`manifest.production.xml`) pointing to the hosted production URL.

## Intended split
- Local/dev manifest: keep localhost URLs for development and sideloading.
- Production manifest: keep `manifest.production.xml` as the publishing manifest for GitHub, Azure, and Microsoft 365 Admin Center rollout.
- Root `manifest.xml`: intentionally removed; only `dist/manifest.xml` is generated during the build.

## Required production preparation before tenant-wide installation
1. Host stable build on Azure Static Web Apps from `main`.
2. Keep `manifest.production.xml` as the source for production rollout.
3. Run the production build; it copies `manifest.production.xml` to `dist/manifest.xml` for the deployed Azure Static Web Apps artifact.
4. Run `npm run check:manifests` to catch local/production URL mixups.
5. Validate the production manifest with `npm run validate:prod` before rollout.
6. Admin-center deployment should use the production manifest only.

## Do not
- Do not deploy `manifest.local.xml` (localhost) to tenant-wide production.
- Do not mix test URLs and production URLs in the same release manifest.

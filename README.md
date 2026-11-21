# Canary Outlook add-in

Authoritative source, tests and offline build tools for the Outlook add-in. Canary CMS consumes a pinned build bundle and continues to host the same paths. There is no hosting migration and no deployment in this change.

## Build and test

Requires Node.js 20 or later; core tests and build have no npm dependencies.

```sh
npm test
CANARY_PUBLIC_URL=https://canary.example.org npm run build
```

`dist/site/` contains the static site fragment. `dist/bundle.json`, `hosting.mjs` and `lock.json` are the CMS build input. The lock records this repository's commit and SHA-256 digests. Build from a clean committed checkout when updating the CMS pin. Without CANARY_PUBLIC_URL the manifest retains its existing template origin, as the CMS did previously.

## Hosting contract

Preserve `/outlook-addin/manifest.xml`, the task/compose/command pages, `/outlook-addin/auth-callback.html`, `/icons/icon{16,32,64,80,128}.png`, and `/.well-known/microsoft-officeaddins-allowed.json` on the existing CMS HTTPS origin. Manifest identity, version 1.0.28.0, permissions, Office.js CDN reference, runtime query versions and callback paths are unchanged. Do not install the same add-in ID at a new origin as part of this source move.

The server's authentication, Graph and mail-plugin APIs, `/connect/mail-plugin` UI and nginx routing remain in canarycms. Shared icon copies are included to make this build self-contained; changes require coordination with CMS branding.

## Update the CMS pin

1. Review source changes and run tests here.
2. Commit, then build from the clean checkout.
3. Copy dist/bundle.json, dist/hosting.mjs and dist/lock.json into the CMS `frontend/vendor/canary-outlook/` directory in a reviewed change. Do not edit generated public/outlook-addin files.
4. Run the CMS frontend build and tests with its existing CANARY_PUBLIC_URL. The adapter verifies the pin, stages assets before Vite, then renders and checks the full hosted bundle after Vite.

Actions remain disabled. No deployment, release upload, Microsoft admin registration or live mailbox testing is automated by this repository. Optional browser fixtures require Playwright and Chromium; see tests/browser.cjs. The source licence remains in SOURCE-LICENSE.txt.

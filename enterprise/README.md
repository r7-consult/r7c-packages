# Enterprise Static Catalog POC

This folder contains local helpers for testing `r7c-packages` as an internal
static catalog.

Russian step-by-step local deployment notes are available in
`enterprise/LOCAL_DEPLOY_RU.md`.

## Start the catalog

```powershell
.\enterprise\serve-static-catalog.ps1 -Port 8090
```

The server exposes the repository root at `http://127.0.0.1:8090/` and adds CORS
headers for browser/plugin-manager requests.

Useful endpoints:

```text
http://127.0.0.1:8090/health
http://127.0.0.1:8090/store/config.json
http://127.0.0.1:8090/sdkjs-plugins/content/hello-world/config.json
```

## Configure r7c

In the `r7c` repository, copy the values from
`enterprise/runtime-config.local.example.js` into
`store/scripts/runtime-config.js`, or pass equivalent URL parameters to
`store/index.html`.

The important value is:

```js
catalogBaseUrl: 'http://127.0.0.1:8090/'
```

Set `enableRatings: false` for offline tests so the UI does not call the public
ONLYOFFICE rating proxy.

## Offline check

After the store UI has been configured to use the local catalog, block access to
GitHub/raw.githubusercontent.com or disconnect external internet access. The
catalog should still load plugin cards, icons, README/License files and plugin
configs from `http://127.0.0.1:8090/`.

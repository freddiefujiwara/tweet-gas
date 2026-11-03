# tweet-gas

A Google Apps Script (GAS) project that posts tweets to X (@freddiefujiwara).

## What it does

- Posts a tweet with the X API (`POST /2/tweets`).
- Signs each request with OAuth 1.0a (HMAC-SHA1). The code makes the signature by itself.
- Only one person can use the web app. The request must have a Google ID token, and the email in the token must be the authorized user's email address.
- After posting, it gets the tweet one more time by ID and checks the text.
- It writes all API responses to the `log` sheet of the spreadsheet.

## Logs

Each API call adds one row to the `log` sheet:

- `time`: when the call happened
- `action`: what happened (`fetch` or `doPost`)
- `detail`: a JSON string with the HTTP method, the URL, the status code, and the full response text

The URL query string is not saved, because the token check URL has the ID token in it. If writing a log fails, posting still works.

## Script Properties

The API keys are not in the code. Set them in the Apps Script project settings.

| Key | Value |
| --- | --- |
| `X_CONSUMER_KEY` | Consumer Key from the X developer console |
| `X_CONSUMER_SECRET` | Consumer Secret from the X developer console |
| `X_ACCESS_TOKEN` | Access Token (Read and write) |
| `X_ACCESS_TOKEN_SECRET` | Access Token Secret |
| `GOOGLE_CLIENT_ID` | (optional) If set, the ID token must be for this client |
| `SPREADSHEET_ID` | (optional, staging) Open this spreadsheet by ID instead of the container spreadsheet |
| `DRY_RUN` | (optional, staging) `1` / `true` = never call the X API; posting returns a synthetic dry-run result (no X credentials needed) |

## Setup

1. Make a GAS project bound to a spreadsheet, and set its Script ID in `.clasp.json`.
2. Set the Script Properties above.
3. Add the `CLASPRC_JSON` secret to this GitHub repository.
4. Add `.github/workflows/node.js.yml` by hand in the GitHub web UI. It is the same workflow as in time-table-gas.
5. Push to main. CI runs the tests, builds the code, pushes it to Apps Script, and makes a new deployment.
6. Copy the web app `/exec` URL to `config.js` in the `tweet` frontend repository.

## Test it

- Run `postTestTweet()` in the Apps Script editor. It posts a test tweet with the current time, and writes the tweet URL to the log.
- Open the `/exec` URL with GET. It returns `{"ok":true,"service":"tweet-gas",...}`.

## Development

```bash
npm install
npm test
npm run build
```

- `npm test` runs the tests. The signature code is tested with the official RFC 5849 test vector.
- `npm run build` makes `dist/Code.gs` from `src/Code.js`.

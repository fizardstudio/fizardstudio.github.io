# Order Sheet Automation (Google Apps Script)

A Google Sheets add-on-style script for small businesses that live in
spreadsheets. It adds a **Data Tools** menu that:

1. **Imports a CSV** from Google Drive into a `Raw` sheet (kept as text, so Sheets
   doesn't silently re-interpret dates)
2. **Cleans and validates** every row: trims and title-cases names, lower-cases
   e-mails, reads 3 date formats (and rejects impossible dates like 31/02),
   strips `$` and thousands separators from prices, maps status aliases
   (`done` -> `shipped`, `canceled` -> `cancelled`)
3. **Removes duplicate order ids** and keeps the first one
4. Writes clean data to `Orders` and puts **every problem row in `Errors`**,
   with the original row number and the exact reason, so a person can fix it
5. Builds a **`Summary`** sheet: revenue by month, product and status, plus a chart
6. **E-mails an HTML report** and can **schedule itself every morning**

**[Watch the 68-second demo video](https://fizardstudio.github.io/video/order-sheet-automation-demo.mp4)**

On the included sample (`sample/orders_raw.csv`, 126 rows) it keeps 115 clean
orders and flags 11 rows for review: 6 duplicate ids, plus one each of invalid
date, invalid e-mail, quantity 0, missing product and unknown status.

## Install (5 minutes)

1. Create a Google Sheet -> **Extensions -> Apps Script**
2. Replace the contents of `Code.gs` with [`AllInOne.gs`](AllInOne.gs)
   (the same code as [`src/Logic.js`](src/Logic.js) + [`src/Code.js`](src/Code.js) in one file)
3. Project Settings -> tick *Show "appsscript.json"* -> paste
   [`src/appsscript.json`](src/appsscript.json) (change `timeZone` to yours)
4. Reload the sheet -> **Data Tools** menu appears -> approve the permissions once

Or with [clasp](https://github.com/google/clasp): `clasp create --type sheets --rootDir src && clasp push`.

Optional: put a report recipient in `Settings!B1` (default is your own address).

## Design notes

* All cleaning rules live in `Logic.js`, which has **no Google API calls**,
  so it is unit-tested in Node:

  ```bash
  node --test
  ```
* Permissions are kept narrow: `spreadsheets.currentonly` and `drive.readonly`
  rather than full Drive access.
* Dates are handled as calendar days, so a report never shifts an order into
  the previous day because of the script's time zone.

All names and e-mails in the sample are fictional.

## License

This code is shown as a portfolio sample. All rights reserved; contact me for licensing. See [LICENSE](../LICENSE).

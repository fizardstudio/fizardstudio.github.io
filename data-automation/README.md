# Data Automation Portfolio - Mochamad Firman

Three small, complete projects showing the kind of work I do for clients:
taking messy business data out of spreadsheets and exports, cleaning it
reliably, and automating the reports people build by hand every week.

| Project | Stack | What it shows |
|---|---|---|
| [Customer Data Cleaner](python-data-cleaning/) | Python, pandas, pytest | Cleans and de-duplicates a messy CRM export - **200,000 rows in under 4 seconds** - with a full audit trail of every change |
| [Order Sheet Automation](google-apps-script/) | Google Apps Script, Google Sheets | Custom menu: CSV import, validation, duplicate removal, summary sheet with chart, daily e-mailed report |
| [Sales Consolidator](excel-vba/) | Excel VBA | Merges monthly sheets or a folder of workbooks into one clean master table, flags bad rows, builds a summary |

## Demo videos

* [Order Sheet Automation in Google Sheets (68 s)](https://fizardstudio.github.io/video/order-sheet-automation-demo.mp4)
* [Sales Consolidator in Excel (46 s)](https://fizardstudio.github.io/video/sales-consolidator-demo.mp4)

## How I approach data work

* **Never silently drop or guess.** Every input row ends up in the clean output,
  a duplicates log, or an issues list with a reason.
* **Rules live in one place**, so they're easy to adapt to a new client's file.
* **Tested.** The Python and Apps Script logic ship with unit tests.
* **Fast on real volumes.** Vectorised pandas and array-based VBA, not cell-by-cell loops.

Background: years of verifying and reconciling land-registry data, plus
building internal business applications (Laravel/PHP, Go, MySQL).

All sample data in this repository is synthetic.

## License

This code is shown as a portfolio sample. All rights reserved; contact me for licensing. See [LICENSE](LICENSE).

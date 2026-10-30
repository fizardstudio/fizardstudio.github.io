# Sales Consolidator (Excel VBA)

The classic "every month someone sends a new sheet" problem, solved with one
button. The macro:

1. **Consolidates** every sheet with the expected header (Jan, Feb, Mar, ...) -
   or every workbook in a folder you pick - into one `Master` table, and records
   the source sheet and row of every line
2. **Cleans** it: trims stray spaces and non-breaking spaces, fixes casing
   (`CARLA DIAZ` -> `Carla Diaz`), converts text dates (`2024-03-05`,
   `05/03/2024`, `5 Mar 2024`) and text numbers (`$1,234.50`, `149 USD`) into
   real Excel values, and adds a `Revenue` column
3. **Flags instead of guessing**: impossible dates (31/02), zero or non-numeric
   units/prices, missing region/product/invoice and **duplicate invoice numbers**
   all go to an `Issues` sheet with the source sheet, row and reason
4. **Summarises** revenue by region and by product, with a chart

**[Watch the 46-second demo video](https://fizardstudio.github.io/video/sales-consolidator-demo.mp4)**

## Try it

1. Open `sales_q1_sample.xlsx` (three months of deliberately messy, fictional data)
2. `Alt+F11` -> **File -> Import File...** -> `SalesConsolidator.bas`
3. Save as **Excel Macro-Enabled Workbook (.xlsm)**
4. `Alt+F8` -> **RunAll** -> Run (or assign `RunAll` to a button)

Other entry points: `ConsolidateSheets`, `ConsolidateFolder` (opens each file
read-only and never modifies it), `CleanMaster`, `BuildSummary`.

## Why it's fast and safe

* Reads and writes whole ranges as arrays - no cell-by-cell loops - and turns
  off screen updating/recalculation while it works, so 100k+ rows stay quick
* No library references to set (late-bound `Scripting.Dictionary`)
* Number parsing does not depend on the PC's regional decimal setting
* Source files are opened read-only, and the original monthly sheets are never changed

Windows Excel 2013 or newer. (`ConsolidateFolder` uses the Windows folder picker.)

## Creating answer forms for menu orders, collecting data from them and generating a result table

### How it works

There are two scripts in `scripts` folder:

`1. raw - answer forms.ps1` - creates many answer forms from `raw.xlsx`. The latter is an excel table in the specified format that contains a list of dishes (it must be called 'raw'). The script collects data from it and transforms to a much simpler and readable table (`sample\sample.xlsx`). Then this table is copied many times to the root folder and every copy is named due to every line in `list.txt` - it contains, obviously, a list of people who are going to make menu orders in this time. This is how we get all of our answer forms.

`2. answer forms - result.ps1` - creates `result.xlsx` from all of the answer forms (they might be filled or not filled out yet). The result table generates a filled due to answers and formatted sheet with menu for each day (these can be instantly printed out without any corrections). It also analyzes answers and makes additional warnings if an order is unusual: combined from two submenu or not complete. The last sheet contains total results for each day: amount of ordered dishes, their prices and summarised costs.

The scripts are written in PowerShell with wide usage of VBA module.

### How to fill out answer forms

An answer form is an excel table named specially for every person from the list. If one wants to order a dish, he/she should highlight the cell containing a dish name with a specific color:
- yellow - 1 unit
- green - 2 units
- purple - 3 units

Pay an attention: these should be 'standard' colors from the excel color palette.

### Additional info

If any error appears when scripts are running, the script is stopped without savings and the error is logged in `ERROR_LOGS.txt` (if not exists, it will be created) in `scripts` folder.

The first script interacts with the next files: `scripts\raw.xlsx`, `scripts\list.txt`, `scripts\sample\sample.xlsx`. It creates/rewrites many answer forms in the root folder.

The second script interacts with the next files: `scripts\json\menu.json`, `scripts\sample\sample.xlsx`, `scripts\list.txt`, many answer forms in the root folder. It creates/rewrites `scripts\result.xlsx`.

All files created/rewrited by scripts must be closed before script running. Otherwise, the error is thrown.

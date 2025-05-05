
# excelextract

**excelextract** is a straightforward command-line tool designed to help you pull structured data out of Excel spreadsheets automatically. You tell it where your data is and what you want using a simple configuration file (written in JSON), and it generates clean CSV files ready for analysis.

It’s particularly helpful for researchers, data collectors, or anyone working with standardized forms, surveys, or logs stored in Excel files.

## What It Does

* Reads data from Excel (.xlsx) files specified by path patterns (including wildcards like `*` and `**`).
* Lets you automatically repeat extraction steps across different sheets, rows, or columns based on your setup.
* Uses a JSON configuration file where you define *exactly* where to find the data and how it should be structured.
* Outputs the extracted data into clean, easy-to-use CSV files.

## Installation

You can install excelextract using `pip`, Python's package installer. Open your terminal or command prompt and run:

```bash
pip install excelextract
```

**Note:** You need Python 3.8 or higher installed on your system to use this tool. This documentation assumes you have Python and `pip` available.

## How It Works: The Configuration File

The core of `excelextract` is the JSON configuration file (e.g., `config.json`). This file tells the tool everything it needs to know. It generally contains:

1.  **Input (`input`): Which Files to Process**

      * Specifies the Excel file(s) to read using a path string.
      * Supports standard file path wildcards (glob patterns like `*`, `?`, `**`) to easily select multiple files (e.g., `"data/*.xlsx"` or `"data/**/*.xlsx"`).
      * Can also be a list of such path strings.

2.  **Output (`output`): Where to Save Results**

      * Defines the name of the output CSV file.

3.  **Lookups (`lookups`): Find Where the Data Is**

      * Lookups are instructions to find specific locations or patterns within the matched Excel files.
      * Each lookup performs a specific **`operation`** (like `looprows`, `findrow`, `loopsheets`).
      * They can find row numbers, column letters, or sheet names matching a pattern.
      * Crucially, lookups can define **Tokens** (like `%%ROW_NUMBER%%` or `%%DATA_SHEET%%`). Think of tokens as named placeholders for values that might change (like the current row number you're processing).

4.  **Columns (`columns`): Define What Data to Extract**

      * This section defines the columns that will appear in your output CSV file.
      * For each output column, you specify its name and how to get its `value` (often using Tokens defined in `lookups`).
      * You specify the data `type` (`string`, `number`) for the output column.
      * You also control *when* a row should be created using **Triggers**.

## Simple Example: Extracting a Single Table

Let's start with a basic task. Imagine you have one Excel file named `report.xlsx` in a folder called `source_data`. Inside this file, there's a sheet named "Data". On this sheet, participant information starts at row 3. You want to extract the 'Name' from column B and 'Value' from column D for all rows where *either* the name *or* the value is present.

Here’s a configuration file (`config.json`) to do this:

```json
{
  "exports": [
    {
      "input": "source_data/report.xlsx",    // Path to the specific input Excel file
      "output": "extracted_data.csv",      // Name of the CSV file to create
      "lookups": [
        {
          "operation": "looprows",           // We need to loop through rows
          "token": "DATA_ROW",             // Create a placeholder named %%DATA_ROW%% for the current row number
          "start": 3,                      // Start checking from row 3
          "count": 500                     // Check up to 500 rows (a safe upper limit)
        }
      ],
      "columns": [
        {
          "name": "participant_name",        // Name of the first column in our output CSV
          "type": "string",                  // Expect text data
          "value": "Data!B%%DATA_ROW%%"      // Get value from Sheet 'Data', Column B, at the current %%DATA_ROW%%
                                             // No trigger specified, so it defaults to "nonempty"
        },
        {
          "name": "measurement_value",     // Name of the second column in our output CSV
          "type": "number",                  // Expect numeric data
          "value": "Data!D%%DATA_ROW%%"      // Get value from Sheet 'Data', Column D, at the current %%DATA_ROW%%
                                             // Also defaults to "nonempty" trigger
        }
      ]
    }
  ]
}
```

**Explanation:**

  * `input`: Specifies the exact path to the single input file.
  * `output`: The name of the CSV file that will be generated.
  * `lookups`: Defines the `%%DATA_ROW%%` token via the `looprows` operation.
  * `columns`: Defines the output columns. Both default to the `nonempty` trigger, so a row is created if either cell B or cell D (in the current `%%DATA_ROW%%`) contains data.

## Advanced Example: Handling Multiple Survey Files Recursively

Now for a more complex, realistic scenario common in research:

  * You have **multiple Excel files** spread across a folder named `data` and potentially its subfolders (e.g., `data/year1/`, `data/year2/`). Each file represents a survey response.
  * Each file has an "overview" sheet with metadata (like the survey name in cell C2).
  * Each file also has **several data sheets** named like "Survey A", "Survey B", etc.
  * On each "Survey X" sheet, there's a table listing participants starting in **column C**, but the exact starting row varies. This header might be "Participants" or "Subjects".
  * You want to extract all participant rows from all "Survey X" sheets across all found files into a **single combined CSV file**. You only want rows where a `participant_id` exists.

Here’s the configuration (`config.json`):

```json
{
  "exports": [
    {
      "input": "data/**/*.xlsx",          // Use glob pattern: find all .xlsx files in 'data' and ALL subdirectories
      "output": "participants_all.csv",   // Combined output file name
      "lookups": [
        {
          "operation": "loopsheets",        // Find sheets within each file based on a pattern
          "token": "SURVEY_SHEET",          // Create placeholder %%SURVEY_SHEET%% for the matching sheet name
          "regex": "Survey .*"              // Pattern: find sheets starting with "Survey " (Note: loopsheets still uses regex for sheet names)
        },
        {
          "operation": "findrow",           // Find the row containing the header text
          "token": "HEADER_ROW",            // Create placeholder %%HEADER_ROW%% for the row number found
          "sheet": "%%SURVEY_SHEET%%",      // Search within the sheet found by loopsheets
          "column": "C",                    // Only search in column C
          "match": ["Participants", "Subjects"], // Find row where cell C contains either "Participants" or "Subjects"
          "select": "first"                 // Take the first match if both appear (or multiple times)
        },
        {
          "operation": "looprows",          // Loop through potential participant rows
          "token": "ROW",                   // Create placeholder %%ROW%% for the current row number being checked
          "start": "%%HEADER_ROW%%",        // Start looping from the header row found above...
          "startOffset": 1,                 // ...but add 1 to start on the row *below* the header
          "count": 100                      // Check up to 100 rows (set a reasonable limit per table)
        }
      ],
      "columns": [
        {
          "name": "source_file",            // Output column for the original Excel filename
          "type": "string",
          "value": "%%FILE_NAME%%",         // Use the special built-in token for the filename
          "trigger": "never"                // Metadata: include this info, but don't use it to decide if a row exists
        },
        {
          "name": "survey_form_name",      // Output column for the survey name from the overview sheet
          "type": "string",
          "value": "overview!C2",           // Get value directly from cell C2 of the 'overview' sheet
          "trigger": "never"                // Metadata: also doesn't trigger row creation
        },
        {
          "name": "participant_id",         // Output column for the participant ID
          "type": "string",
          "value": "%%SURVEY_SHEET%%!B%%ROW%%", // Get value from the current survey sheet, column B, current row
          "trigger": "nonempty"             // CRITICAL: Explicitly require this field to be non-empty to create a row
        },
        {
          "name": "participant_score",      // Output column for the participant score
          "type": "number",
          "value": "%%SURVEY_SHEET%%!I%%ROW%%", // Get value from the current survey sheet, column I, current row
          "trigger": "nonempty"              // Explicitly set trigger (same as default) - row created if this OR participant_id is non-empty
        }
      ]
    }
  ]
}
```

**Explanation:**

  * `input: "data/**/*.xlsx"`: Uses `**` to find files recursively.
  * `operation`: Replaces `type` in the `lookups`.
  * `match`: Used in `findrow` to look for either "Participants" or "Subjects" without using regex syntax.
  * `trigger`: The comment on `participant_score` is corrected. Setting `trigger: "nonempty"` explicitly here has the same effect as the default, meaning a row is created if *either* `participant_id` (which also has `nonempty`) OR `participant_score` has data in the source cell for that `%%ROW%%`.

## Running the Tool

Once you have your configuration file (e.g., `config.json`) ready:

1.  Open your terminal or command prompt.
2.  Navigate to the directory where your `config.json` file is saved.
3.  Run the tool by typing:

<!-- end list -->

```bash
excelextract config.json
```

This will execute the instructions in `config.json` and generate the specified CSV output file(s) in the same directory.

## Configuration Reference

This section provides details on all available options for your JSON configuration file.

Each configuration file must have a top-level `exports` key, which contains a list `[...]` of one or more "export jobs". Each job defines one process of reading from input file(s) and writing to an output file.

```json
{
  "exports": [
    {
      "input": "path/pattern/or/list/of/patterns", // e.g., "data/*.xlsx" or ["data/a*.xlsx", "data/b*.xlsx"]
      "output": "output_filename.csv",
      "lookups": [ /* List of lookup operations */ ],
      "columns": [ /* List of column definitions */ ]
    }
    // You can add more export jobs here if needed
  ]
}
```

Note: all keys/fields are case insensitive.

### Top-Level Fields per Export Job

| Field    | Type             | Required | Description                                                                                                                                                              |
|----------|------------------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `input`  | String or List   | Yes      | Specifies the input Excel file(s). <br> - If String: A path potentially containing glob wildcards (`*`, `?`, `**`). <br> - If List: A list of path strings (each supporting glob). |
| `output` | String           | Yes      | The name of the CSV file to be created with the extracted data.                                                                                                            |
| `lookups`| List             | Yes      | A list of lookup operations to find locations and define tokens. Executed in order.                                                                                        |
| `columns`| List             | Yes      | A list defining the columns for your output CSV file.                                                                                                                    |

*(Note: Glob patterns like `*` match any characters except path separators, `?` matches a single character, `**` matches directories recursively.)*

### Lookup Operations (`lookups`)

Lookups find locations or iterate over parts of your Excel files, often defining **Tokens** (placeholders like `%%NAME%%`) that you can reuse later.

**Common Fields for all Lookups:**

| Field       | Type   | Description                                                                         |
|-------------|--------|-------------------------------------------------------------------------------------|
| `operation` | String | The kind of lookup operation (e.g., `"looprows"`, `"findrow"`).                       |
| `token`     | String | The name for the placeholder (Token) this lookup defines (e.g., `"ROW_NUM"`).         |

**Lookup Operation: `loopsheets`**

Loops over sheets in the current workbook that match a regex pattern.

| Field       | Type   | Description                                                                               |
|-------------|--------|-------------------------------------------------------------------------------------------|
| `operation` | String | Must be `"loopsheets"`.                                                                   |
| `token`     | String | Name of the token that will hold the matching sheet name (e.g., `"DATA_SHEET"`).           |
| `regex`     | String | The pattern (regex) to match sheet names against (e.g., `"Survey.*"` matches Survey1, Survey2, ...). |

**Lookup Operations: `findRow`, `findColumn`**

Finds the row number(s) or column letter(s) containing specific text. Useful for locating headers or markers whose position might change or have alternative names.

| Field       | Type              | Description                                                                                                                                                                                                    |
|-------------|-------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `operation` | String            | Must be `"findRow"` or `"findColumn"`.                                                                                                                                                                         |
| `token`     | String            | Name of the token that will hold the row number(s) or column letter(s) found (e.g., `"HEADER_ROW"`). If multiple values are selected (via `select`), the token might hold a list.                               |
| `sheet`     | String            | Name of the sheet to search within. Can use tokens (e.g., `"%%DATA_SHEET%%"`). Required.                                                                                                                       |
| `match`     | String or List    | Required. Specifies the text to find in cells. <br> - If String: Matches cells containing this exact (case-sensitive) string. <br> - If List: Matches cells containing exactly (case-sensitive) any string in the list. |
| `column`    | String            | **For `findRow` only:** Optional. Restrict the search to this specific column (e.g., `"A"`).                                                                                                                  |
| `row`       | Number            | **For `findColumn` only:** Optional. Restrict the search to this specific row number (e.g., `1`).                                                                                                               |
| `select`    | String/Num/List | Optional (default: `"first"`). Controls which match(es) to use if multiple cells match the criteria. <br> - `"first"`: Use the first match found. <br> - `"last"`: Use the last match found. <br> - *Future:* Integer or List support. |
| `unique` | Bool | Optional (default: `false`). Regardless which match is selected (e.g., `"first"`), if this is set and more than 1 match is found, the export exists with an error |

**Lookup Operation: `findcell`**

Searches an entire sheet for a cell containing specific text and defines tokens for both its row number and column letter. This is useful when you need to locate a specific anchor point in a sheet and then potentially use its row and column context for further extractions.

| Field       | Type   | Description                                                                                                                                                                                                     |
|-------------|--------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `operation`   | String           | Must be `"findcell"` |
| `match`     | String or List | The exact (case-sensitive) text to find within a cell on the specified sheet.                                                                                                                                                                                           |
| `sheet`     | String | Name of the sheet to search within. Can use tokens (e.g., `"%%DATA_SHEET%%"`).                                                                                                                                                                                           |
| `rowtoken`  | String | Name of the token that will hold the row number of the found cell (e.g., `"HEADER_ROW"`).                                                                                                                                                                                          |
| `columntoken` | String | Name of the token that will hold the column letter of the found cell (e.g., `"DATA_COLUMN"`).                                                                                                                                                                                      |
| `unique` | Bool | Optional (default: `false`). If this is set and more than 1 match is found, the export exists with an error |

This operation is the exception on the rule, which does not have the `token` field, but generates 2 tokens: `rowtoken` and `columntoken`.

**Lookup Operations: `loopRows`, `loopColumns`**

Iterates through a range of row numbers or column letters, assigning the current value to a token for each step.

| Field         | Type             | Description                                                                                                                          |
|---------------|------------------|--------------------------------------------------------------------------------------------------------------------------------------|
| `operation`   | String           | Must be `"loopRows"` or `"loopColumns"`.                                                                                               |
| `token`       | String           | Name of the token that will hold the current row number or column letter (e.g., `"ROW"`, `"COL"`).                                       |
| `start`       | String or Number | Starting row number or column letter. Can be a literal (`1`, `"B"`) or a token (`"%%HEADER_ROW%%"`). Required.                          |
| `end`         | String or Number | Optional. The row number or column letter to stop at (inclusive). Cannot be used together with `count`.                                  |
| `count`       | Number           | Optional. The maximum number of steps to take. Useful if you don't know the exact end but want to set a limit. Cannot be used with `end`. |
| `startOffset` | Number           | Optional (default 0). Adds this offset to the `start` value. E.g., `start: "%%HEADER_ROW%%", startOffset: 1` starts one row *below* the header. |
| `endOffset`   | Number           | Optional (default 0). Adds this offset to the `end` value.                                                                           |
| `stride`      | Number           | Optional (default 1). Step size for the loop. E.g., `stride: 2` processes every second row/column.                                   |

### Column Definitions (`columns`)

Defines the structure of your output CSV file. Each object in the `columns` list corresponds to one column in the CSV.

| Field       | Type    | Required | Description                                                                                                                                                                                             |
|-------------|---------|----------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `name`      | String  | Yes      | The header name for this column in the output CSV file.                                                                                                                                                 |
| `type`      | String  | Yes      | Data type for the output column. Use `"string"` for text, `"number"` for numeric values (integers or decimals). Affects how data is read and potentially formatted.                                        |
| `value`     | String  | Yes      | How to get the value for this column. Can be: <br> - A **literal string**: `value: "Constant Value"` <br> - A **cell reference**: `value: "SheetName!A1"` or using tokens: `value: "%%SHEET%%!B%%ROW%%"` <br> - A **token**: `value: "%%FILE_NAME%%"` |
| `trigger`   | String  | No       | Controls if this column can trigger the creation of a new row in the CSV. Options: `"nonempty"` (default), `"never"`, `"nonzero"`. See Trigger System below.                                           |
| `rowOffset` | Number  | No       | Optional (default 0). Adds an offset to the row number part of a cell reference in `value`. Useful for getting data from adjacent rows (e.g., `value: "Data!A%%ROW%%", rowOffset: 1` gets data from row below). |
| `colOffset` | Number  | No       | Optional (default 0). Adds an offset to the column part of a cell reference in `value`. Useful for getting data from adjacent columns.                                                                  |

**Important `value` Syntax:**

  * If `value` contains an exclamation mark (`!`), it's treated as a **cell reference** (`SheetName!CellAddress`). Tokens like `%%ROW%%` or `%%SHEET%%` will be substituted before reading the cell.
  * If `value` *only* contains a **token** (e.g., `%%FILE_NAME%%`), the token's value is used directly.
  * Otherwise, `value` is treated as a **literal string** that will be put directly into the CSV cell.

### Trigger System (`trigger`)

The `trigger` property on a column definition controls *if* and *when* a new row is added to your output CSV file.

  * If the `trigger` key is **not specified** for a column, it **defaults to `"nonempty"`**.
  * For each potential data point identified by your `lookups` (e.g., for each row number in a `looprows`), the tool checks the `trigger` conditions of your columns.
  * If **at least one** column definition has its trigger condition met for the current context (e.g., current `%%ROW%%`), a new row is created in the CSV file.
  * Once a row is created, the tool calculates the `value` for **all** defined columns for that specific context and writes them to the CSV. If a referenced cell for a column is empty, an empty value will be written in the CSV for that column in the created row.

**Available Trigger Values:**

  * `"nonempty"` (Default):
      * The trigger condition is met if the cell referenced in the `value` field is **not empty**.
      * Use this (or rely on the default) for key data columns that indicate a valid record exists.
  * `"never"`:
      * This column **never** triggers the creation of a new row, even if its referenced cell has data.
      * Use this for metadata columns (like `source_file` or fixed survey details) that you want to include *alongside* the main data but shouldn't determine if a row exists on their own. If all data columns are empty but a metadata column with `trigger: "never"` references a non-empty cell, no row will be created.
  * `"nonzero"`:
      * The trigger condition is met only if the cell referenced in the `value` field contains a **numeric value** that is **not equal to zero**. Blank cells or cells with text do not meet this condition.
      * Useful if you only want to include rows where a specific measurement or count is actually greater than zero.

## Built-in Tokens

`excelextract` provides one special token that is always available:

  * `%%FILE_NAME%%`: Holds the filename (including extension but excluding the path) of the Excel file currently being processed.

Other tokens (like `%%ROW%%`, `%%SURVEY_SHEET%%`, `%%HEADER_ROW%%`) are defined by you within the `lookups` section using the `token` field.

## Key Features Summary

  * Extracts data based on a clear JSON configuration, separating settings from the tool itself.
  * Selects input files using intuitive path strings with glob pattern support (`*`, `**`, `?`), accepting a single path or a list of paths.
  * Uses **Tokens** (like `%%ROW%%`, `%%SURVEY_SHEET%%`) as placeholders for dynamic values like row numbers or sheet names found during processing.
  * Supports various lookup **operations** (`loopsheets`, `looprows`, `loopcolumns`, `findrow`, `findcolumn`) to locate data dynamically.
  * Can find specific rows/columns based on cell content using `match` (for exact or alternative strings) without requiring regex, while `loopsheets` still uses regex for sheet name patterns. Uses `select` to choose which match(es) to use.
  * Flexible **Trigger System** (defaulting to `nonempty`) to precisely control when data rows are created based on whether key cells contain data or meet specific conditions (`nonzero`).
  * Combines data from multiple sheets and multiple Excel files into single CSV outputs.
  * Outputs standard CSV files (UTF-8 encoded) compatible with most data analysis tools and spreadsheets.

## License

MIT License © 2025 Philippe

> This project is not affiliated with or endorsed by Microsoft. "Excel" is a registered trademark of Microsoft Corporation. This tool uses the .xlsx format purely as a data source.

## Contributing

This tool is shared in the hope it helps others with structured data collection workflows. Pull requests, feedback, and improvements are welcome! Please feel free to open an issue or submit a pull request on the project repository.

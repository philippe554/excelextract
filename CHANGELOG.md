# Changelog

All notable changes to `excelextract` will be documented in this file.  

## [Unreleased]

## [0.3.0] - 2025-05-05

- Add findcell which searches for both a row and a column in a full sheet
- Add unique property which fails the export if more than 1 match is found
- Make config file keys case insensitive

## [0.2.0] - 2025-05-05

- Introduced Core Extraction Logic: The package now provides the fundamental ability to read data from .xlsx files based on a user-defined configuration.
- JSON Configuration: Implemented support for defining extraction rules using a JSON configuration file (config.json).
- Input File Selection: Added functionality to specify input Excel files using file paths with support for glob patterns (*, ?, **) for selecting multiple files.
- Lookup Operations: Introduced various lookup operations (loopsheets, findRow, findColumn, loopRows, loopColumns) to dynamically locate data within workbooks and sheets.
- Token System: Added support for defining and using tokens (e.g., %%ROW%%, %%SHEET_NAME%%) as placeholders for dynamic values found during lookups.
- Cell Content Matching: Implemented the findRow and findColumn operations with a match property for finding cells based on specific text content (exact string or list of alternatives) and a select property for handling multiple matches.
- Configurable Row Triggering: Developed a trigger system (defaulting to nonempty, with options for never and nonzero) on column definitions to control when a new row is created in the output CSV based on cell content.
- Multi-File and Multi-Sheet Extraction: Enabled combining data from multiple sheets and multiple input files into a single output CSV.
- Built-in %%FILE_NAME%% Token: Included a pre-defined token to easily reference the name of the currently processed input file.
- CSV Output: Added the capability to output extracted data into standard, UTF-8 encoded CSV files.

## [0.1.0] – 2025-05-03

Initial release on PyPI.
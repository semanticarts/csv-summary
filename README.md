# csv-summary
Generates a basic analysis of a CSV to an XLSX workbook.

```
usage: csv_summary [-h] [-o OUTPUT_PATH]
                   [--category-threshold CATEGORY_THRESHOLD]
                   [--date-format DATE_FORMAT]
                   [--date-time-format DATE_TIME_FORMAT] [-i IGNORE_VALUE]
                   [--sheet-name SHEET_NAME] [-s NUM_SAMPLES]
                   input

positional arguments:
  input                 CSV or XSL input containing data. If XLS/XLSX, will
                        add two extra sheets with summary/sample data. Will
                        update in place unless -o is specified. First sheet is
                        treated as the data unless --sheet-name is specified.

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT_PATH, --output_path OUTPUT_PATH
                        Path to output .xlsx, default is same directory/name
                        as CSV input
  --category-threshold CATEGORY_THRESHOLD
                        Columns that have equal or less unique values will be
                        treated as a category, and their values will be
                        counted and output in the summary.
  --date-format DATE_FORMAT
                        Regular expression for detecting date/time columns.
  --date-time-format DATE_TIME_FORMAT
                        Regular expression for detecting date/time columns.
  -i IGNORE_VALUE, --ignore-value IGNORE_VALUE
                        Ignore these values from summary, use for blank
                        equivalents such as '?' and 'N/A'
  --sheet-name SHEET_NAME
                        Process XLS with the input in the specified sheet. If
                        -o is not specified, will add new sheets to input XLS.
  -s NUM_SAMPLES, --num-samples NUM_SAMPLES
                        Number of rows to sample in transposed view
```

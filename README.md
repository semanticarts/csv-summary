# csv-summary
Generates a basic analysis of a CSV to an XLSX workbook.

```
usage: csv_summary [-h] [-o OUTPUT_PATH]
                   [--category-threshold CATEGORY_THRESHOLD]
                   [--date-format DATE_FORMAT] [-i IGNORE_VALUE]
                   [-s NUM_SAMPLES]
                   csv_path

positional arguments:
  csv_path

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
  -i IGNORE_VALUE, --ignore-value IGNORE_VALUE
                        Ignore these values from summary, use for blank
                        equivalents such as '?' and 'N/A'
  -s NUM_SAMPLES, --num-samples NUM_SAMPLES
                        Number of rows to sample in transposed view
```

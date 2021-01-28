# CSV Summary Tool

import csv
import re
from argparse import ArgumentParser, FileType
from collections import defaultdict
from os.path import splitext

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


def parse_arguments():
    parser = ArgumentParser()
    parser.add_argument("csv_path", type=FileType("r", encoding="utf-8"))
    parser.add_argument("-o", "--output_path",
                        help="Path to output .xlsx, default is same directory/name as CSV input",
                        type=FileType("w"))
    parser.add_argument("--category-threshold", type=int, default=100,
                        help="Columns that have equal or less unique values will be treated as a category, "
                             "and their values will be counted and output in the summary.")
    parser.add_argument("--date-format", action="store",
                        default="\\d{2}[-/]\\d{2}[-/]\\d{4}[- ]\\d{2}:\\d{2}:\\d{2}(\\.\\d{1,6})?",
                        help="Regular expression for detecting date/time columns.")
    parser.add_argument("-i", "--ignore-value", action="append", default=[],
                        help="Ignore these values from summary, use for blank equivalents such as '?' and 'N/A'")
    parser.add_argument("-s", "--num-samples", type=int, default=3,
                        help="Number of rows to sample in transposed view")

    return parser.parse_args()


def summarize_csv():
    args = parse_arguments()
    output_path = args.output_path
    if not output_path:
        (base, ext) = splitext(args.csv_path.name)
        output_path = open(base + ".xlsx", "wb")
    wb = Workbook()

    reader = csv.reader(args.csv_path)
    headers = next(reader)
    by_column = dict((header, defaultdict(int)) for header in headers)
    all_dates = defaultdict(lambda: True)

    csv_copy = wb.active
    csv_copy.title = "Data"
    csv_copy.append(headers)

    date_regex = re.compile(args.date_format)

    num_rows = 0
    for line in reader:
        csv_copy.append(list("" if val in args.ignore_value else val for val in line))
        for col, value in zip(headers, line):
            if value not in args.ignore_value:
                by_column[col][value] += 1
                all_dates[col] &= date_regex.fullmatch(value) is not None

        num_rows += 1

    header_row(csv_copy)
    auto_width(csv_copy)

    summary = wb.create_sheet("Summary")
    summary.append(headers)
    for index, header in enumerate(headers):
        values = by_column[header]
        if len(values) == num_rows:
            # All values are unique
            summary.cell(row=2, column=index + 1).value = "Unique"
        elif len(values) > 0 and all_dates[header]:
            # All values are dates
            summary.cell(row=2, column=index + 1).value = "Dates"
        elif len(values) <= args.category_threshold:
            for row, usage in enumerate(
                    f"{k} [{v}]" for k, v in sorted(values.items(), reverse=True, key=lambda item: item[1])):
                summary.cell(row=row + 2, column=index + 1).value = usage
    header_row(summary)
    auto_width(summary)

    samples = wb.create_sheet("Samples")
    for row, hdr in enumerate(headers):
        samples.cell(row=row+1, column=1).value = hdr
    for sample in range(args.num_samples):
        for index in range(len(headers)):
            samples.cell(row=index + 1, column=sample + 2).value = csv_copy.cell(row=sample + 2, column=index + 1).value
    header_col(samples)
    auto_width(samples)

    wb.save(output_path)


def auto_width(sheet):
    column_widths = [0] * len(next(sheet.iter_rows()))
    for row in sheet.iter_rows():
        for i, cell in enumerate(row):
            column_widths[i] = max(column_widths[i], len(str(cell.value)))
    for i, column_width in enumerate(column_widths):
        sheet.column_dimensions[get_column_letter(i + 1)].width = column_width


def header_row(sheet):
    bold = Font(bold=True)
    for col in range(len(next(sheet.iter_rows()))):
        sheet.cell(column=col + 1, row=1).font = bold
    sheet.freeze_panes = 'A2'


def header_col(sheet):
    bold = Font(bold=True)
    for row in range(len(sheet['A'])):
        sheet.cell(column=1, row=row + 1).font = bold
    sheet.freeze_panes = 'B1'


if __name__ == '__main__':
    summarize_csv()

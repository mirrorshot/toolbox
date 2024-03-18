"""
requirements:
msoffcrypto-tool
openpyxl
"""

import argparse
import datetime
import io
import json
import re
from collections import ChainMap
from itertools import islice
from typing import Any

import msoffcrypto
import openpyxl
from openpyxl.cell import Cell

version = "1.0.0"


def get_argument_parser(
    parent_parser: argparse.ArgumentParser | None = None,
) -> argparse.ArgumentParser:
    arg_parser = parent_parser or argparse.ArgumentParser(
        description="Tool used to extract data from an Excel spreadsheet to a json file."
    )
    arg_parser.add_argument("file_path", type=str, help="Excel spreadsheet path")
    arg_parser.add_argument(
        "-p",
        "--password",
        type=str,
        help="Password if the given xls/xlsx spreadsheet encrypted is encrypted.",
        required=False,
        default=None,
    )
    default_sheet_name = "Sheet1"
    arg_parser.add_argument(
        "-s",
        "--sheet",
        type=str,
        help=f'Name of the xls/xlsx sheet where the data are stored. Default value = "{default_sheet_name}"',
        required=False,
        default=default_sheet_name,
    )
    default_output_file = "newsletter-user-data.json"
    arg_parser.add_argument(
        "-o",
        "--output",
        type=str,
        help=f'Output file path. Default value = "{default_output_file}"',
        required=False,
        default=default_output_file,
    )
    arg_parser.add_argument(
        "-P",
        "--pretty",
        type=bool,
        help="When active, the output json will be printed with a 2 character indentation.",
        required=False,
        default=False,
        action=argparse.BooleanOptionalAction,
    )
    arg_parser.add_argument(
        "-f",
        "--format-file",
        metavar="format_file",
        type=str,
        help="""Used to provide a custom mapping format defined in a json file.
        Example: {
          "row": "{_row_number}",
          "email": "{email}",
          "user": {
            "name": "{name}",
            "surname": "{surname}"
          },
          "operation_datetime": "{_now}"
        }
        The "{label}" notation searches for a field with that name in the header row.
        The following fields are runtime generated:
        - _row_number
        - _now
        Fields mapping supports the standard Python format notation for formatted strings.
        """,
        required=False,
        default=None,
    )
    default_mapping_labels = ["id", "name", "surname", "email", "created_at"]
    arg_parser.add_argument(
        "-m",
        "--mapping-labels",
        metavar="mapping_labels",
        type=str,
        help=f'Used to provide a custom mapping labels list. Default value = "{default_mapping_labels}"',
        nargs="+",
        required=False,
        default=default_mapping_labels,
    )
    default_unique_key = "email"
    arg_parser.add_argument(
        "-u",
        "--unique-key",
        metavar="unique_key",
        help=f'Used to provide the column number or label for the unique key. If not provided, the column with id "{default_unique_key}" will be used.',
        required=False,
        default=default_unique_key,
    )
    return arg_parser


def decrypt(file_path: str, password: str) -> io.BytesIO:
    decrypted_workbook = io.BytesIO()

    with open(file_path, "rb") as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_workbook)

    return decrypted_workbook


def load_workbook(file_path: str, password: str | None = None) -> openpyxl.Workbook:
    return openpyxl.load_workbook(
        file_path if password is None else decrypt(file_path, password)
    )


def find_column_index(row: tuple[Cell, ...], label: str) -> int:
    return next(i for i, cell in enumerate(row) if cell.value == label)


def map_data(
    workbook: openpyxl.Workbook,
    sheet: str,
    mapping_labels: list[str],
    unique_key: int | str | None,
) -> list[dict[str, Any]]:
    def value(cell: Cell):
        return (
            cell.value
            if type(cell.value) is not datetime.datetime
            else cell.value.strftime("%Y-%m-%dT%H:%M:%S") + ".000Z"
        )

    def map_row(row: tuple[Cell, ...]) -> dict:
        return {v: value(row[i]) for i, v in enumerate(mapping_labels)}

    unique_key_index = (
        unique_key
        if type(unique_key) is int
        else find_column_index(row=next(workbook[sheet].rows), label=unique_key)
    )
    return list(
        {
            row[unique_key_index].value: map_row(row)
            for row in islice(workbook[sheet].rows, 1, None)
        }.values()
    )


def load_format(format_file: str) -> dict:
    with open(format_file, "r") as file:
        return json.load(file)


def map_formatted_data(
    workbook: openpyxl.Workbook,
    sheet: str,
    mapping_format: dict,
    unique_key: int | str | None,
) -> list[dict]:
    _now = datetime.datetime.now()

    def clean_label(label):
        match = re.match(pattern=r"{(\w*)(:.*)?}", string=label)
        # print(f"{label = } -> {match.groups() = }")
        return match.group(1)

    def extract_labels(mapping: dict[str, Any]) -> list[str]:
        labels = []
        for v in mapping.values():
            if type(v) is str:
                labels.append(clean_label(v))
            else:
                labels.extend(extract_labels(v))
        return labels

    mapped_labels = extract_labels(mapping_format)
    label_indexes = {
        h.value: i
        for i, h in enumerate(next(workbook[sheet].rows))
        if h.value in mapped_labels
    }
    # print(f"{mapped_labels = }")
    # print(f"{label_indexes = }")

    def _map_formatted(row: dict[str, Any], mf: dict | str) -> dict | str:
        return (
            mf.format(**row)
            if type(mf) is str
            else {k: _map_formatted(row, v) for k, v in mf.items()}
        )

    def map_formatted_row(row_number: int, row: tuple[Cell, ...], mf: dict) -> dict:
        dict_row = {label: row[index].value for label, index in label_indexes.items()}
        # print(f"{dict_row = }")
        return _map_formatted(
            dict(
                ChainMap(
                    {"_row_number": row_number, "_now": _now},
                    dict_row,
                )
            ),
            mf,
        )

    unique_key_index = (
        unique_key if type(unique_key) is int else label_indexes[unique_key]
    )
    return list(
        {
            row[unique_key_index].value: map_formatted_row(i, row, mapping_format)
            for i, row in islice(enumerate(workbook[sheet].rows), 1, None)
        }.values()
    )


def save_json(file_path: str, data: list[dict[str, Any]], pretty: bool) -> None:
    with open(file_path, "w") as f:
        f.write(json.dumps(data, indent=2 if pretty else None))


def main(
    file_path: str,
    sheet: str,
    password: str,
    output: str,
    pretty: bool,
    mapping_labels: list[str],
    unique_key: int | str,
    format_file: str | None,
) -> None:
    workbook = load_workbook(file_path=file_path, password=password)

    if format_file is None:
        data = map_data(
            workbook=workbook,
            sheet=sheet,
            mapping_labels=mapping_labels,
            unique_key=unique_key,
        )
    else:
        data = map_formatted_data(
            workbook=workbook,
            sheet=sheet,
            mapping_format=load_format(format_file=format_file),
            unique_key=unique_key,
        )

    save_json(file_path=output, data=data, pretty=pretty)


if __name__ == "__main__":
    argument_parser = get_argument_parser()
    argument_parser.add_argument(
        "-v", "--version", action="version", version=f"%(prog)s {version}"
    )
    args = argument_parser.parse_args()
    main(**vars(args))

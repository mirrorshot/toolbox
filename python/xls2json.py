"""
requirements:
msoffcrypto-tool
openpyxl
"""

import argparse
import datetime
import io
import json
from itertools import islice
from typing import Any

import msoffcrypto
import openpyxl
from openpyxl.cell import Cell

version = "1.0.0"


def get_argument_parser(
        parent_parser: argparse.ArgumentParser | None = None
) -> argparse.ArgumentParser:
    arg_parser = parent_parser or argparse.ArgumentParser(
        description="Tool used to extract data from an Excel spreadsheet to a json file."
    )
    arg_parser.add_argument(
        "file_path",
        type=str,
        help="Excel spreadsheet path"
    )
    arg_parser.add_argument(
        "-p", "--password",
        type=str,
        help="Password if the given xls/xlsx spreadsheet encrypted is encrypted.",
        required=False,
        default=None,
    )
    default_sheet_name = "Sheet1"
    arg_parser.add_argument(
        "-s", "--sheet",
        type=str,
        help=f"Name of the xls/xlsx sheet where the data are stored. Default value = \"{default_sheet_name}\"",
        required=False,
        default=default_sheet_name
    )
    default_output_file = "newsletter-user-data.json"
    arg_parser.add_argument(
        "-o", "--output",
        type=str,
        help=f"Output file path. Default value = \"{default_output_file}\"",
        required=False,
        default=default_output_file
    )
    arg_parser.add_argument(
        "-P", "--pretty",
        type=bool,
        help="When active, the output json will be printed with a 2 character indentation.",
        required=False,
        default=False,
        action=argparse.BooleanOptionalAction,
    )
    default_mapping_labels = ["id", "name", "surname", "email", "created_at"]
    arg_parser.add_argument(
        "-m", "--mapping_labels",
        type=str,
        help=f"Used to provide a custom mapping labels list. Default value = \"{default_mapping_labels}\"",
        nargs="+",
        required=False,
        default=default_mapping_labels,
    )
    default_unique_key = "email"
    arg_parser.add_argument(
        "-u", "--unique_key",
        help=f"Used to provide the column number or label for the unique key. If not provided, the column with id \"{default_unique_key}\" will be used.",
        required=False,
        default=default_unique_key,
    )
    return arg_parser


def decrypt(file_path: str, password: str) -> io.BytesIO:
    decrypted_workbook = io.BytesIO()

    with open(file_path, 'rb') as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_workbook)

    return decrypted_workbook


def load_workbook(
        file_path: str,
        password: str | None = None
) -> openpyxl.Workbook:
    return openpyxl.load_workbook(
        file_path
        if password is None
        else decrypt(file_path, password)
    )


def find_column_index(
        row: tuple[Cell, ...],
        label: str
) -> int:
    return next(i for i, cell in enumerate(row) if cell.value == label)


def map_data(
        workbook: openpyxl.Workbook,
        sheet: str,
        mapping_labels: list[str],
        unique_key: int | str | None = None,
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
    return list({row[unique_key_index].value: map_row(row) for row in islice(workbook[sheet].rows, 1, None)}.values())


def save_json(
        file_path: str,
        data: list[dict[str, Any]],
        pretty: bool
) -> None:
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
) -> None:
    workbook = load_workbook(file_path=file_path, password=password)
    data = map_data(workbook=workbook, sheet=sheet, mapping_labels=mapping_labels, unique_key=unique_key)
    save_json(file_path=output, data=data, pretty=pretty)


if __name__ == '__main__':
    argument_parser = get_argument_parser()
    argument_parser.add_argument(
        "-v", "--version",
        action="version",
        version=f"%(prog)s {version}"
    )
    args = argument_parser.parse_args()
    main(**vars(args))

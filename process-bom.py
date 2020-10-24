#!/usr/bin/python3

from colorama import Fore, Style
from collections import namedtuple
from openpyxl import load_workbook
from pathlib import Path
import csv

BomLine = namedtuple("BomLine", "val footprint mpn")


class SubFinderCsv:
    "Locate generic part numbers using CSV"

    def __init__(self, csv_path):
        with open(csv_path, newline="") as f:
            self.subs = list(
                csv.DictReader(
                    f, fieldnames=["value", "footprint", "mpn"], delimiter=","
                )
            )

    def find(self, part):
        for row in self.subs:
            if (
                part.val.value == row["value"]
                and part.footprint.value == row["footprint"]
            ):
                return row["mpn"]
        print(
            f"{Fore.YELLOW}No part number found for {Style.BRIGHT}{part.val.value},{part.footprint.value}{Style.RESET_ALL}{Fore.RESET}"
        )

        return ""


class BomXlsGen:
    "Iterable interface to Xls Bom"

    def __init__(self, wb):
        self.wb = wb

        sheet = wb["BoM"]
        self.r = sheet.rows

        # Skip to data headers, and find MPN column
        for row in self.r:
            if row[0].value == "Row":
                # Locate and record MPN column index
                for cell in row:
                    if cell.value == "MPN":
                        self.mpn_idx = cell.column - 1
                        break
                break

    @classmethod
    def from_path(cls, path):
        return cls(load_workbook(filename=path))

    def __iter__(self):
        return self

    def __next__(self):
        for row in self.r:
            # Return pertinent parts of BoM line
            return BomLine(row[3], row[5], row[self.mpn_idx])
        raise StopIteration()

    def save(self, path):
        self.wb.save(path)


def update_bom(in_path, out_path, subs_path):

    bom = BomXlsGen.from_path(in_path)
    generic_finder = SubFinderCsv(subs_path)

    for line in bom:
        if line.mpn.value == "":
            line.mpn.value = generic_finder.find(line)

    bom.save(out_path)


if __name__ == "__main__":
    from argparse import ArgumentParser

    parser = ArgumentParser(
        description="Update missing MPN lines from substitution list based on value and footprint"
    )
    parser.add_argument("input", type=Path, help="Input BOM file")
    parser.add_argument("output", type=Path, help="Output BOM file")
    parser.add_argument("subs_csv", type=Path, help="List of substitutions in csv")

    args = parser.parse_args()

    update_bom(args.input, args.output, args.subs_csv)

#!/usr/bin/env python3

import logging
import argparse
import pandas as pd
import csv
from datetime import datetime
import re
import io
from dateparser import parse

DATECOL_REGEX = re.compile(r"date|time", re.IGNORECASE)
DATE_STRING = r"%m/%d/%Y %H:%M:%S %p"

ILLEGAL_CHARACTERS_RE = re.compile("[\\000-\\010]|[\\013-\\014]|[\\016-\\037]")

LOGGER = logging.getLogger("csv2xlsx")
LOGGER.setLevel(logging.INFO)

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")


def parsedate(s, fmt):
    ret = None
    try:
        if s:
            if fmt:
                ret = datetime.strptime(s, fmt)
            else:
                ret = parse(s)
        else:
            LOGGER.warn(f"Unable to parse {s} as date")
    finally:
        return ret


def main():

    parser = argparse.ArgumentParser(
        description="Take CSV as input and emit xlsx with fixed dates"
    )

    parser.add_argument("infile", metavar="INFILE.csv", help="Input CSV file", nargs=1)

    parser.add_argument("outfile", metavar="OUTFILE.xlsx", help="Outfile file", nargs=1)

    parser.add_argument(
        "-f", "--date-fields", nargs="*", dest="datecols", help="list of date fields"
    )

    parser.add_argument(
        "--date-string",
        help="format string, defaults to guess",
        default=None,
        dest="datefmt",
    )

    parser.add_argument(
        "--skip-rows", default=0, type=int, help="skip N rows", dest="skiprows"
    )
    parser.add_argument(
        "--dont-autodetect-date-fields",
        action="store_false",
        dest="autodetect_date",
        default=True,
        help="Do not autodetect date fields",
    )

    parser.add_argument(
        "--debug", "-d", action="store_true", help="Debug level messages"
    )

    args = parser.parse_args()
    if args.debug:
        LOGGER.setLevel(logging.DEBUG)

    buf = None
    with open(args.infile[0], "rb") as f:
        blob = f.read().decode("utf8", "ignore")
        LOGGER.debug("Removing bad characters")
        fixed_blob = re.sub(ILLEGAL_CHARACTERS_RE, "", blob)
        buf = io.StringIO(fixed_blob)

    df = pd.read_csv(buf, skiprows=args.skiprows, encoding="utf8")  # args.infile[0],
    r, c = df.shape
    LOGGER.debug(f"Succesfully read {r} rows from {args.infile[0]}")

    datecols = []
    if args.datecols:
        LOGGER.debug(f"Additional date cols: {args.datecols}")
        datecols.extend(args.datecols)
    if args.autodetect_date:
        for c in df.columns:
            if DATECOL_REGEX.search(c):
                LOGGER.debug(f"Detected {c} as date column")
                datecols.append(c)

    for c in datecols:
        LOGGER.debug(f"Processing column {c}")
        def _parsedate(s):
            d=parsedate(s, args.datefmt)
            if d is not None:
                # Excel does not understand timezone aware formats
                return d.replace(tzinfo=None)
            return d
        df[c] = df[c].apply(_parsedate)

    with pd.ExcelWriter(args.outfile[0]) as excel:
        df.to_excel(excel, index=False)


if __name__ == "__main__":
    main()

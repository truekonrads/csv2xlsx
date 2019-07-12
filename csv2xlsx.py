#!/usr/bin/env 

import logging
import argparse
import pandas as pd
import csv
from datetime import datetime
import re

DATECOL_REGEX=re.compile(r"date",re.IGNORECASE)
DATE_STRING=r"%m/%d/%Y %H:%M:%S %p"
#CHANGME: Add the name of the logger - e.g. your program name
LOGGER = logging.getLogger("csv2xlsx")
LOGGER.setLevel(logging.INFO)

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

def parsedate(s,fmt):
    ret=None
    try:
        if s:
            ret=datetime.strptime(s,fmt)
        else:
            LOGGER.warn(f"Unable to parse {s} as date")
    finally:
        return ret
def main():

    #CHANGME: Describe what your program does
    parser = argparse.ArgumentParser(
        description='Take CSV as input and emit xlsx with fixed dates')

    parser.add_argument("infile",
            metavar="INFILE.csv",
            help="Input CSV file",
            nargs=1)

    parser.add_argument("outfile",
            metavar="OUTFILE.xlsx",
            help="Outfile file",
            nargs=1)

    parser.add_argument(
        "--date-fields",
        "-f",
        nargs="*",
        dest="datecols",
        help="list of date fields")

    parser.add_argument(
        "--date-string",
        default=DATE_STRING,
        help=f"format string, default: {DATE_STRING}",
        dest="datefmt")

    parser.add_argument(
        "--dont-autodetect-date-fields",
        action="store_false",
        dest="autodetect_date",
        default=True,
        help="Do not autodetect date fields")

    parser.add_argument('--debug', "-d", 
            action="store_true",
            help='Debug level messages')
    
    args = parser.parse_args()
    if args.debug:
        LOGGER.setLevel(logging.DEBUG)

    
    df=pd.read_csv(args.infile[0])
    r,c = df.shape
    LOGGER.debug(f"Succesfully read {r} rows from {args.infile[0]}")

    datecols=[]
    if args.datecols:
        LOGGER.debug(f"Additional date cols: {args.datecols}")
        datecols.extend(args.datecols)
    if args.autodetect_date:        
        for c in df.columns:
            if (DATECOL_REGEX.search(c)):
                LOGGER.debug(f"Detected {c} as date column")
                datecols.append(c)   
    

    for c in datecols:
        LOGGER.debug(f"Processing column {c}")
        df[c]=df[c].apply(lambda s: parsedate(s,args.datefmt))    
    with pd.ExcelWriter(args.outfile[0]) as excel:
        df.to_excel(excel,index=False)
if __name__ == '__main__':
    main()

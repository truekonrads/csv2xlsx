# csv2xlsx
Take CSV and convert to xlsx while fixing date formats

```
usage: csv2xlsx.py [-h] [--date-fields [DATECOLS [DATECOLS ...]]]
                   [--date-string DATEFMT] [--dont-autodetect-date-fields]
                   [--debug]
                   INFILE.csv OUTFILE.xlsx

Take CSV as input and emit xlsx with fixed dates

positional arguments:
  INFILE.csv            Input CSV file
  OUTFILE.xlsx          Outfile file

optional arguments:
  -h, --help            show this help message and exit
  --date-fields [DATECOLS [DATECOLS ...]], -f [DATECOLS [DATECOLS ...]]
                        list of date fields
  --date-string DATEFMT
                        format string, default: %m/%d/%Y %H:%M:%S %p
  --dont-autodetect-date-fields
                        Do not autodetect date fields
  --debug, -d           Debug level messages
  ```

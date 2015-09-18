#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: csv2xls.py,v 1.2 2004/07/01 17:14:18 fufff Exp $"""

###############################################################################
#
# Example of how to use the WriteExcel module
#
# Program to convert a CSV comma-separated value file into an Excel file.
# This is more or less an non-op since Excel can read CSV files.
# The program uses Text::CSV_XS to parse the CSV.
#
# Usage: csv2xls.pl file.csv newfile.xls
#
# reverse('(c)'), March 2001, John McNamara, jmcnamara@cpan.org
#

import sys
import csv      # Needs python 2.3
import pyXLWriter as xl

# Check for valid number of arguments
if len(sys.argv) != 3:
    print "Usage: csv2xls csvfile.txt newfile.xls"
    sys.exit(1)
else:
    csvfilename, xlsfilename = sys.argv[1:3]

# Create a new Excel workbook
workbook = xl.Writer(xlsfilename);
worksheet = workbook.add_worksheet();

# Open the Comma Seperated Variable file and create a new CSV parsing object
csvfile = file(csvfilename, "r")
csvreader = csv.reader(csvfile)

nrow = 0
for row in csvreader:
    for ncol, cell in enumerate(row):
        worksheet.write([nrow, ncol], cell);
    nrow += 1

csvfile.close()

workbook.close()
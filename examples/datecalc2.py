#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: datecalc2.py,v 1.3 2004/07/01 17:14:18 fufff Exp $"""

###############################################################################
#
# Example of how to using the Date::Calc module to calculate Excel dates.
#
# See also the excel_date1.pl example.
#
# These functions have been superceded by Spreadsheet::WriteExcel::Utility.
#
# reverse('(c)'), June 2001, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl
from datetime import datetime       # Needs python 2.3


def excel_date(date):
    """Create an Excel date in the 1900 format. All of the arguments are optional
    but you should at least add years.

    Corrects for Excel's missing leap day in 1900. See excel_time1.pl for an
    explanation.
    
    """
    epoch = datetime(1899, 12, 31, 0, 0, 0)
    delta = date - epoch
    xldate = delta.days + float(delta.seconds) / (24*60*60)
    # Add a day for Excel's missing leap day in 1900
    if xldate > 59:
        xldate += 1
    return xldate


def excel_date_1904(date):
    """Create an Excel date in the 1904 format. All of the arguments are optional
    but you should at least add years.
    
    You will also need to call workbook.set_1904() for this format to be valid.
    
    """
    epoch = datetime(1904, 1, 1, 0, 0, 0)
    delta = date - epoch
    xldate = delta.days + float(delta.seconds) / (24*60*60)
    return xldate
    
    
def main():
    """Simple example."""
    
    # Create a new workbook and add a worksheet
    workbook = xl.Writer("excel_date2.xls")
    worksheet = workbook.add_worksheet()
    
    # Expand the first column so that the date is visible.
    worksheet.set_column("A:A", 25)
    
    # Add a format for the date
    format =  workbook.add_format()
    format.set_num_format("d mmmm yyy HH:MM:SS")
    
    # Write some dates and times
    date =  excel_date(datetime(1900, 1, 1))
    worksheet.write("A1", date, format)
    
    date = excel_date(datetime(2000, 1, 1))
    worksheet.write("A2", date, format)
    
    date = excel_date(datetime(2000, 4, 17, 14, 33, 15))
    worksheet.write("A3", date, format)

    workbook.close()


if __name__ == "__main__":
    main()
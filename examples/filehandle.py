#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: filehandle.py,v 1.5 2004/02/08 14:47:03 fufff Exp $"""

###############################################################################
#
# Example of using Spreadsheet::WriteExcel to write to alternative filehandles.
#
# reverse('(c)'), April 2003, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl
from StringIO import StringIO

###############################################################################
#
# Example 1. This demonstrates the standard way of creating an Excel file by
# specifying a file name.
#
workbook1  = xl.Writer('fh_01.xls')
worksheet1 = workbook1.add_worksheet()

worksheet1.write([0, 0],  "Hi Excel!")
workbook1.close()

###############################################################################
#
# Example 2'. Write an Excel file to an existing filehandle.
#
test = file("fh_02_.xls", "wb")

workbook2  = xl.Writer(test)
worksheet2 = workbook2.add_worksheet()

worksheet2.write([0, 0],  "Hi Excel!")
workbook2.close()

###############################################################################
#
# Example 4'. Write an Excel file to a string via IO::Scalar. Please refer to
# the IO::Scalar documentation for further details.
# (Python's analog is StringIO/cStringIO)
xls = StringIO()

workbook4  = xl.Writer(xls)
worksheet4 = workbook4.add_worksheet()

worksheet4.write([0, 0], "Hi Excel 4")
workbook4.close()       # This is required before we use the scalar

tmp = file("fh_04_.xls", "wb")
tmp.write(xls.getvalue())
tmp.close()

###############################################################################
#
# Example 4''. Write an Excel file to a string via IO::Scalar. Please refer to
# the IO::Scalar documentation for further details.
# (Python's analog is StringIO/cStringIO)
xls = StringIO()

workbook4  = xl.Writer(xls, big=True)
worksheet4 = workbook4.add_worksheet()

worksheet4.write([0, 0], "Hi Excel 4")
workbook4.close()       # This is required before we use the scalar

tmp = file("fh_04_2.xls", "wb")
tmp.write(xls.getvalue())
tmp.close()
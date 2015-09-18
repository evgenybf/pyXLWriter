#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: cgi.py,v 1.1 2004/01/31 18:57:53 fufff Exp $"""

###############################################################################
#
# Example of how to use the Spreadsheet::WriteExcel module to send an Excel
# file to a browser in a CGI program.
#
# On Windows the hash-bang line should be something like:
# #!C:\Perl\bin\perl.exe
#
# reverse('(c)'), March 2001, John McNamara, jmcnamara@cpan.org
#

import sys
import pyXLWriter as xl

# Set the filename and send the content type
filename = "cgitest.xls"

print "Content-type: application/vnd.ms-excel"
# The Content-Disposition will generate a prompt to save the file. If you want
# to stream the file to the browser, comment out the following line.
print "Content-Disposition: attachment; filename=%s" % filename
print 

# Create a new workbook and add a worksheet. The special Perl filehandle - will
# redirect the output to STDOUT
#
workbook = xl.Writer(sys.stdout)
worksheet = workbook.add_worksheet()

# Set the column width for column 1
worksheet.set_column(0, 20)

# Create a format
format = workbook.add_format()
format.set_bold();
format.set_size(15);
format.set_color('blue')

# Write to the workbook
worksheet.write([0, 0], "Hi Excel!", format)

workbook.close()
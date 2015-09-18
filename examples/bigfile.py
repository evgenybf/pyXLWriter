#!/usr/bin/python -u
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: bigfile.py,v 1.6 2004/02/22 18:29:40 fufff Exp $"""

########################################################################
#
# Example of how to extend the Spreadsheet::WriteExcel 7MB limit with
# OLE::Storage_Lite: http://search.cpan.org/search?dist=OLE-Storage_Lite
#
# Nov 2000, Kawai, Takanori (Hippo2000)
#   Mail: GCD00051@nifty.ne.jp
#   http://member.nifty.ne.jp/hippo2000
#

from time import time
import pyXLWriter as xl


########################################################################
#
# Manually flag (using OLEWriterBig)
#
workbook1  = xl.Writer("small_big.xls", big=True)
worksheet = workbook1.add_worksheet()

worksheet.write("A1", "Used OLEWriterBig!")

workbook1.close()


########################################################################
#
# big=True is automatically, because, file is too big
#
t = time()
colcount = 100 + 1
rowcount = 6000 + 1

workbook2  = xl.Writer("big.xls") 
worksheet = workbook2.add_worksheet()

worksheet.set_column((0, colcount), 18)

print "Filling..."
for col in xrange(colcount):
    print "[%d]" % col, 
    for row in xrange(rowcount):
        worksheet.write_string((row, col), "Row: %d Col: %d" % (row, col))
print "ok!"

workbook2.close()

print "Time: %.2f" % (time() - t)
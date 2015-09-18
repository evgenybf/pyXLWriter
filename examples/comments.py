#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: comments.py,v 1.9 2004/01/31 18:56:07 fufff Exp $"""

###############################################################################
#
# This example demonstrates writing cell comments. A cell comment is indicated
# in Excel by a small red triangle in the upper right-hand corner of the cell.
#
# reverse('(c)'), April 2003, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl

workbook = xl.Writer("comments.xls")
worksheet = workbook.add_worksheet()

worksheet.write([2, 2], "Hello")
worksheet.write_comment([2, 2], "This is a comment.")

workbook.close()
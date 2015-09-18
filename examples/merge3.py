#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: merge3.py,v 1.3 2004/01/31 18:56:07 fufff Exp $"""

###############################################################################
#
# Example of how to use Spreadsheet::WriteExcel to write a hyperlink in a
# merged cell. There are two options write_url_range() with a standard merge
# format or merge_range().
#
# reverse('(c)'), September 2002, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl

# Create a new workbook and add a worksheet
workbook = xl.Writer('merge3.xls')
worksheet = workbook.add_worksheet()

# Increase the cell size of the merged cells to highlight the formatting.
for i in (1, 3, 6, 7):
    worksheet.set_row(i, 30) 
worksheet.set_column('B:D', 20)

###############################################################################
#
# Example 1: Merge cells containing a hyperlink using write_url_range()
# and the standard Excel 5+ merge property.
#
format1 = workbook.add_format(merge = 1,
                              border = 1,
                              underline = 1,
                              color = 'blue'
                             )

# Write the cells to be merged
worksheet.write_url_range('B2:D2', 'http://www.perl.com', format=format1)
worksheet.write_blank('C2', format1)
worksheet.write_blank('D2', format1)

###############################################################################
#
# Example 2: Merge cells containing a hyperlink using merge_range().
#
format2 = workbook.add_format(border = 1,
                              underline = 1,
                              color = 'blue',
                              align = 'center',
                              valign = 'vcenter'
                             )

# Merge 3 cells
worksheet.merge_range('B4:D4', 'http://www.perl.com', format2)

# Merge 3 cells over two rows
worksheet.merge_range('B7:D8', 'http://www.perl.com', format2)

workbook.close()
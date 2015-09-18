#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: demo.py,v 1.11 2004/01/31 18:56:07 fufff Exp $"""

#######################################################################
#
# Demo of some of the features of Spreadsheet::WriteExcel.
# Used to create the project screenshot for Freshmeat.
#
#
# reverse('(c)'), October 2001, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl

workbook = xl.Writer("demo.xls")
worksheet = workbook.add_worksheet('Demo')
worksheet2 = workbook.add_worksheet('Another sheet')
worksheet3 = workbook.add_worksheet('And another')

#######################################################################
#
# Write a general heading
#
worksheet.set_column('A:B', 32)
heading  = workbook.add_format(bold = 1,
                               color = 'blue',
                               size = 18,
                               merge = 1,
                              )

headings = ('Features of Spreadsheet::WriteExcel', '')
worksheet.write_row('A1', headings, heading)

#######################################################################
#
# Some text examples
#
text_format = workbook.add_format(bold = 1,
                                   italic = 1,
                                   color = 'red',
                                   size = 18,
                                   font = 'Lucida Calligraphy'
                                  )

worksheet.write('A2', "Text")
worksheet.write('B2', "Hello Excel")
worksheet.write('A3', "Formatted text")
worksheet.write('B3', "Hello Excel", text_format)

#######################################################################
#
# Some numeric examples
#
num1_format = workbook.add_format(num_format='#,##0.00')
num2_format = workbook.add_format(num_format=' d mmmm yyy')

worksheet.write('A4', "Numbers")
worksheet.write('B4', 1234.56)
worksheet.write('A5', "Formatted numbers")
worksheet.write('B5', 1234.56, num1_format)
worksheet.write('A6', "Formatted numbers")
worksheet.write('B6', 37257, num2_format)

#######################################################################
#
# Formulae
#
worksheet.set_selection('B7')
worksheet.write('A7', 'Formulas and functions, "=SIN(PI()/4)"')
worksheet.write('B7', '=SIN(PI()/4)')

#######################################################################
#
# Hyperlinks
#
worksheet.write('A8', "Hyperlinks")
worksheet.write('B8',  'http://www.perl.com/' )

#######################################################################
#
# Images
#
worksheet.write('A9', "Images")
worksheet.insert_bitmap('B9', 'republic.bmp', 16, 8)

#######################################################################
#
# Misc
#
worksheet.write('A17', "Page/printer setup")
worksheet.write('A18', "Multiple worksheets")

workbook.close()
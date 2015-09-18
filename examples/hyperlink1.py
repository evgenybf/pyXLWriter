#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: hyperlink1.py,v 1.4 2004/02/23 10:08:37 fufff Exp $"""

###############################################################################
#
# Example of how to use the WriteExcel module to write hyperlinks
#
# See also hyperlink2.pl for worksheet URL examples.
#
# reverse('(c)'), March 2001, John McNamara, jmcnamara@cpan.org
#

import pyXLWriter as xl

# Create a new workbook and add a worksheet
workbook  = xl.Writer("hyperlink.xls")
worksheet = workbook.add_worksheet('Hyperlinks')

# Format the first column
worksheet.set_column('A:A', 30)
worksheet.set_selection('B1')


# Add a sample format
format = workbook.add_format()
format.set_size(12)
format.set_bold()
format.set_color('red')
format.set_underline()


# Write some hyperlinks
worksheet.write('A1', 'http://www.perl.com/'                )
worksheet.write_url('A3', 'http://www.perl.com/', 'Perl home')
worksheet.write('A5', 'http://www.perl.com/', format=format)
worksheet.write_url('A7', 'mailto:jmcnamara@cpan.org', 'Mail me')

# Write a URL that isn't a hyperlink
worksheet.write_string('A9', 'http://www.perl.com/')

workbook.close()
#!/usr/bin/env python
# This example script was ported from Perl Spreadsheet::WriteExcel module.
# The author of the Spreadsheet::WriteExcel module is John McNamara
# <jmcnamara@cpan.org>

__revision__ = """$Id: easter_egg.py,v 1.1 2004/02/01 09:14:27 fufff Exp $"""

###############################################################################
#
# This uses the Win32::OLE module to expose the Flight Simulator easter egg
# in Excel 97 SR2. A must see.
#
# reverse('(c)'), March 2001, John McNamara, jmcnamara@cpan.org
#

import win32com.client.dynamic

application = win32com.client.dynamic.Dispatch("Excel.Application")
workbook = application.Workbooks.Add()
worksheet = workbook.Worksheets(1)

application.Visible = 1

worksheet.Range("L97:X97").Select()
worksheet.Range("M97").Activate()

message = "Hold down Shift and Ctrl and click the " + \
          "Chart Wizard icon on the toolbar.\n\n" + \
          "Use the mouse motion and buttons to control " + \
          "movement. Try to find the monolith. " + \
          "Close this dialog first."

application.InputBox(message)
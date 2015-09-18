#!/usr/bin/env python
__revision__ = """$Id: test_utilites.py,v 1.13 2004/02/12 16:40:37 fufff Exp $"""

import unittest

import testsupport
from pyXLWriter.utilites import *


class UtilitesTest(unittest.TestCase):

    def test_cell_to_rowcol(self):
        rc = cell_to_rowcol2("A1")
        self.assertEqual((0, 0), rc)
        rc = cell_to_rowcol2("E100")
        self.assertEqual((99, 4), rc)
        rc = cell_to_rowcol("A1")
        self.assertEqual((0, 0, False, False), rc)
        rc = cell_to_rowcol("E$100")
        self.assertEqual((99, 4, True, False), rc)

    def test_cellrange_to_rowcol_pair(self):
        rcs = cellrange_to_rowcol_pair("A1")
        self.assertEqual((0, 0, 0, 0), rcs)
        rcs = cellrange_to_rowcol_pair("B7")
        self.assertEqual((6, 1, 6, 1), rcs)
        rcs = cellrange_to_rowcol_pair("4:6")
        self.assertEqual((3, 0, 5, -1), rcs)
        rcs = cellrange_to_rowcol_pair("D:G")
        self.assertEqual((0, 3, -1, 6), rcs)
        rcs = cellrange_to_rowcol_pair("D4:G6")
        self.assertEqual((3, 3, 5, 6), rcs)


if __name__ == "__main__":
    unittest.main()

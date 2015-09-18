#!/usr/bin/env python
__revision__ = """$Id: test_XLWriter.py,v 1.8 2004/02/08 14:47:03 fufff Exp $"""

"""test.test_XLWriter

Python and perl results compare.

See: perl_xl/gen_*.pl and WriteExcelTest's methods test_*

"""

import os, os.path
import unittest

import testsupport
import pyXLWriter as xl


class XLWriterTest(unittest.TestCase):

    #~ def setUp(self):
        #~ pass

    #~ def tearDown(self):
        #~ pass
        
    def test_gen_version(self):
        vi = (0, 3, 2, "alpha", 1)
        v = xl._gen_version(vi)
        self.assertEqual("0.3.2a1", v)
        vi = (0, 3, 0, "alpha", 1)
        v = xl._gen_version(vi)
        self.assertEqual("0.3a1", v)
        vi = (0, 3, 0, "final", 1)
        v = xl._gen_version(vi)
        self.assertEqual("0.3", v)
        vi = (1, 3, 6, "final", 1)
        v = xl._gen_version(vi)
        self.assertEqual("1.3.6", v)


if __name__ == "__main__":
    unittest.main()

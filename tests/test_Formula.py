#!/usr/bin/env python
"""tests.test_Formula

"""
__revision__ = """$Id: test_Formula.py,v 1.7 2004/02/12 07:08:21 fufff Exp $"""

import unittest

import testsupport
from pyXLWriter.Formula import Formula

# TODO: res = self.formula.parse_formula("='List 1'!A1")


class FormulaTest(unittest.TestCase):
    """Test for Formula class"""

    def setUp(self):
        self.formula = Formula()

    def tearDown(self):
        self.formula = None

    def test_not_operations(self):
        res = self.formula.parse_formula("=100")
        etalon = "\x1E\x64\x00"
        self.assertEqual(etalon, res)
        res = self.formula.parse_formula("=6.7")
        etalon = "\x1F\xCD\xCC\xCC\xCC\xCC\xCC\x1A\x40"
        self.assertEqual(etalon, res)
        
    def test_operations(self):
        res = self.formula.parse_formula("=2+3")
        etalon = "\x1E\x02\x00\x1E\x03\x00\x03"
        self.assertEqual(etalon, res)

    def test_ref(self):
        res = self.formula.parse_formula("=A1")
        etalon = "D\x00\xC0\x00"
        self.assertEqual(etalon, res)
        res = self.formula.parse_formula("=A1:D3")
        etalon = "E\x00\xc0\x02\xc0\x00\x03"
        self.assertEqual(etalon, res)
        
    def test_functions(self):    
        res = self.formula.parse_formula("=NOW() * 10")
        etalon = "\x19\x01\x00\x00\x41\x4A\x00\x1E\x0A\x00\x05"
        self.assertEqual(etalon, res)
        res = self.formula.parse_formula("=SUM(10,20,30,40)")
        etalon = "\x1E\x0A\x00\x1E\x14\x00\x1E\x1E\x00\x1E\x28\x00\x42\x04\x04\x00"
        self.assertEqual(etalon, res)
        
    def test_function_with_string(self):
        res = self.formula.parse_formula('IF(A5>3,"Yes", "No")')
        etalon = 'D\x04\xc0\x00\x1e\x03\x00\r\x17\x03Yes\x17\x02NoB\x03\x01\x00'
        self.assertEqual(etalon, res)


if __name__ == "__main__":
    unittest.main()

#!/usr/bin/env python
"""tests.test_parse

"""
__revision__ = """$Id: test_parse.py,v 1.9 2004/03/30 05:19:35 fufff Exp $"""

import unittest

import testsupport
from pyXLWriter.parse import parse


class ParserTest(unittest.TestCase):

    def test_simple(self):
        res = parse("3")
        self.assertEqual(['_num', 3, '_arg', 1], res)
        res = parse('"6" & "s"')
        self.assertEqual(['_str', '"6"', '_str', '"s"', 'ptgConcat', '_arg', 1], res)
        res = parse("-3+7")
        self.assertEqual(['_num', -3, '_num', 7, 'ptgAdd', '_arg', 1], res)
        res = parse("1*2+3")
        self.assertEqual(['_num', 1, '_num', 2, 'ptgMul', '_num', 3, 'ptgAdd', '_arg', 1], res)
        res = parse("1+2*3")
        self.assertEqual(['_num', 1, '_num', 2, '_num', 3, 'ptgMul', 'ptgAdd', '_arg', 1], res)

    def test_ref_and_range(self):
        res = parse("A1")
        self.assertEqual(['_ref2d', 'A1', '_arg', 1], res)

    def test_comma(self):
        res = parse("1+2,3+4")
        self.assertEqual(['_num', 1, '_num', 2, 'ptgAdd', '_num', 3, '_num', 4, 'ptgAdd', '_arg', 2], res)
        res = parse("1*2,3*4")
        self.assertEqual(['_num', 1, '_num', 2, 'ptgMul', '_num', 3, '_num', 4, 'ptgMul', '_arg', 2], res)

    def test_function(self):
        res = parse("SUM(10)")
        self.assertEqual(['_class', 'SUM', '_num', 10, '_arg', 1, '_func', 'SUM', '_arg', 1], res)
        res = parse("NOW() * 20")
        self.assertEqual(['_class', 'NOW', '_func', 'NOW', '_num', 20, 'ptgMul', '_arg', 1], res)
        res = parse("SUM(SUM(B4:G10),A3)+3")
        etalon = [
            "_class", "SUM",
                "_class", "SUM", "_range2d", "B4:G10", "_arg", 1, "_func", "SUM",
                "_ref2d", "A3", "_arg", 2, "_func", "SUM",
            "_num", 3, "ptgAdd", "_arg", 1,]
        self.assertEqual(etalon, res)

    def test_string_in_function(self):
        res = parse('IF(A5>3,"Yes", "No")')
        etalon = ['_class', 'IF', '_ref2d', 'A5', '_num', 3, 'ptgGT', '_str', '"Yes"', '_str', '"No"', '_arg', 3, '_func', 'IF', '_arg', 1]
        self.assertEqual(etalon, res)
        
    def test_gele(self):
        res = parse("6>=6<7=8>=9")
        self.assertEqual(["_num", 6, "_num", 6, "ptgGE", "_num", 7, "ptgLT", "_num", 8, "ptgEQ", "_num", 9, "ptgGE", "_arg", 1], res)

    def test_not_stripped_branches(self):
        res = parse("A1/(A2-A3)")
        self.assertEqual(["_ref2d", "A1", "_ref2d", "A2", "_ref2d", "A3", "ptgSub", "_arg", 1, "ptgParen", "ptgDiv", "_arg", 1], res)


if __name__ == "__main__":
    unittest.main()
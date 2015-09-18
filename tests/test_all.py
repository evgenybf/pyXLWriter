#!/usr/bin/env python
__revision__ = """$Id: test_all.py,v 1.7 2004/01/31 12:31:37 fufff Exp $"""

import unittest
import glob
import os


def suite():
    """Return a test suite containing all test cases in all test modules.
    Searches the current directory for any modules matching test_*.py.

    """
    suite = unittest.TestSuite()
    for filename in glob.glob('test_*.py'):
        module = __import__(os.path.splitext(filename)[0])
        suite.addTest(unittest.defaultTestLoader.loadTestsFromModule(module))
    return suite


if __name__ == '__main__':
    unittest.main(defaultTest='suite')

#!/urs/bin/env python
"""setup

Package installer for pyXLWriter.

"""
__revision__ = """$Id: setup.py,v 1.26 2004/08/20 05:16:16 fufff Exp $"""

import sys
from distutils.core import setup
import pyXLWriter as xl

if sys.version < '2.2.3':
    from distutils.dist import DistributionMetadata
    DistributionMetadata.classifiers = None
    DistributionMetadata.download_url = None

setup(name = "pyXLWriter",
    version = xl.__version__,
    author = "Evgeny Filatov",
    author_email = "fufff@users.sourceforge.net",
    url = "http://sourceforge.net/projects/pyxlwriter/",
    download_url="http://sourceforge.net/project/showfiles.php?group_id=99682&package_id=106979",    
    description = "A library for generating Excel compatible Spreadsheets",
    long_description = """\
pyXLWriter is a Python library for generating Excel compatible spreadsheets.
It's a port of John McNamara's Perl Spreadsheet::WriteExcel module
version 1.01 to Python. It allows writing of Excel spreadsheets
without the need for COM objects.
""",
    license = "GNU LGPL",
    platforms = "Platform Independent",
    packages = ["pyXLWriter"],
    keywords = "xls excel spreadsheet workbook database table",
    classifiers=["Operating System :: OS Independent",
                 "Programming Language :: Python",
                 "License :: OSI Approved :: GNU Library or Lesser General Public License (LGPL)",
                 "Development Status :: 3 - Alpha",
                 "Intended Audience :: Developers",
                 "Topic :: Software Development :: Libraries :: Python Modules",
                 "Topic :: Office/Business :: Financial :: Spreadsheet",
                 "Topic :: Database",
                 "Topic :: Internet :: WWW/HTTP :: Dynamic Content :: CGI Tools/Libraries",
                ],
    )

__revision__ = """$Id: testsupport.py,v 1.7 2004/01/31 12:31:37 fufff Exp $"""

"""test.testsupport

Several functions used in tests.

"""

if 1:
    # I don't like install module for testing.
    import sys
    sys.path.insert(0, "..")


def read_file(filename):
    """Read and return file contents by his filename."""
    f = file(filename, "rb")
    try:
        c = f.read()
    finally:
        f.close()
    return c


def is_equal_files(filename1, filename2):
    """Return True if files are equals, else - False."""
    f1 = read_file(filename1)
    f2 = read_file(filename2)
    return f1 == f2

# nadia/excel.py

#
# Copyright (c) 2010-2011 Simone Basso <bassosimone@gmail.com>
#
# Permission to use, copy, modify, and distribute this software for any
# purpose with or without fee is hereby granted, provided that the above
# copyright notice and this permission notice appear in all copies.
#
# THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
# WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
# MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
# ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
# WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
# ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
# OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
#

#
# This is the foundation: code to read excel files.  We use xlrd
# and xlwt for that.  See <http://www.python-excel.org/>.
#

import pprint
import sys

#
# Try to import and use system-wide xlrd and xlwt, if they are
# available.  In case of failure, fall-back with a private stripped-
# down copy of these two libraries included with Nadia.
#

PREFIX = "@PREFIX@/share/nadia"
if "@PREFIX@" in PREFIX:
    PREFIX = "."

try:
    import xlrd, xlwt.Utils
except ImportError:
    sys.path.insert(0, PREFIX)
    import xlrd, xlwt.Utils

def _convert_range(range):
    v = range.split(":")
    if len(v) != 2:
        raise ValueError("Passed an invalid A1 range")

    firstrow, firstcol = xlwt.Utils.cell_to_rowcol2(v[0])
    lastrow, lastcol = xlwt.Utils.cell_to_rowcol2(v[1])

    # sanity
    if (firstrow < 0 or firstcol < 0 or lastrow < 0 or lastcol < 0 or
      firstrow > lastrow or firstcol > lastcol):
        raise ValueError("Passed an invalid A1 range")

    return firstrow, firstcol, lastrow, lastcol

def open_sheet(path):
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)
    return sheet

def row_values(sheet, range):
    M = []
    firstrow, firstcol, lastrow, lastcol = _convert_range(range)
    row = firstrow
    while row <= lastrow:
        R = sheet.row_values(row, firstcol, lastcol + 1)
        R = map(unicode, R)
        M.append(R)
        row = row + 1
    return M

def cell_value(sheet, cell):
    row, col = xlwt.Utils.cell_to_rowcol2(cell)
    value = sheet.cell_value(row, col)
    value = unicode(value)
    return value

if __name__ == "__main__":
    sheet = open_sheet("test/sample1.xls")
    data = row_values(sheet, "C3:R24")
    pprint.pprint(data)
    cell = cell_value(sheet, "A11")
    pprint.pprint(cell)

# nadia/data.py

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
# Pass 1.  Read the data section.  The data section is composed of two
# matrices: one that contains the headers and the other that contains the
# actual data.  The code in this file compact the headers using a simple
# algorithm that reduces multiple lines of headers to a single line.  Then
# it attaches the data and the result is a data where the first line is
# headers and the remainder is data.
#

import pprint
import sys

if __name__ == "__main__":
    sys.path.insert(0, ".")

from nadia.excel import open_sheet
from nadia.excel import row_values

def _edit_inplace(cells):
    nrows = len(cells)
    ncols = len(cells[0])
    for j in range(1, ncols):
        if not cells[0][j]:
            cells[0][j] = cells[0][j-1]
    for i in range(1, nrows):
        for j in range(1, ncols):
            if not cells[i][j] and cells[i-1][j-1] == cells[i-1][j]:
                cells[i][j] = cells[i][j-1]

def _compress(cells, separator):
    nrows = len(cells)
    ncols = len(cells[0])
    for j in range(0, ncols):
        vector = [ cells[0][j] ]
        for i in range(1, nrows):
            if cells[i][j]:
                vector.append(separator)
                vector.append(cells[i][j])
        heading = u"".join(vector)
        heading = heading.lower()
        cells[0][j] = heading

def _data_headers(sheet, range, separator=u": "):
    cells = row_values(sheet, range)
    _edit_inplace(cells)
    _compress(cells, separator)
    return cells[0]

def _data_values(sheet, values_range):
    cells = row_values(sheet, values_range)
    nrows = len(cells)
    ncols = len(cells[0])
    for j in range(0, ncols):
        for i in range(0, nrows):
            if cells[i][j] == u"1.0":
                cells[i][j] = u"si"
            cells[i][j] = cells[i][j].lower()
    return cells

def data_section(sheet, hdrs_range, values_range):
    hdrs = _data_headers(sheet, hdrs_range)
    values = _data_values(sheet, values_range)

    for row in values:
        while len(row) < len(hdrs):
            row.append(u"")
        while len(row) > len(hdrs):
            hdrs.append(u"")

    data = []
    data.append(hdrs)
    for row in values:
        data.append(row)

    return data

if __name__ == "__main__":
    sheet = open_sheet("test/sample1.xls")
    data = data_section(sheet, "C3:R6", "C7:R24")
    pprint.pprint(data)

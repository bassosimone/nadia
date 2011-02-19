# nadia/finalize.py

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
# This is the last stage of processing.  Here we remove empty lines
# from the data and then we add global data and global constants to
# the data.  Note that we keep using unicode.
#

import pprint
import sys

if __name__ == "__main__":
    sys.path.insert(0, ".")

from nadia.excel import open_sheet
from nadia.excel import row_values
from nadia.excel import cell_value
from nadia.data import data_section
from nadia.yaxis import y_axis

def remove_empty_lines(data):
    ndata = []
    for row in data:
        count = 0
        for item in row:
            count += len(item)
        if count == 0:
            continue
        ndata.append(row)
    return ndata

def global_data(data, sheet, name, cell):
    data[0].append(unicode(name))
    value = cell_value(sheet, cell)
    for row in data[1:]:
        row.append(value)

def global_const(data, name, value):
    data[0].append(unicode(name))
    for row in data[1:]:
        row.append(unicode(value))

if __name__ == "__main__":
    sheet = open_sheet("test/sample1.xls")
    data = data_section(sheet, "C3:R6", "C7:R24")
    y_axis(sheet, "B7:B24", data)
    data = remove_empty_lines(data)
    global_data(data, sheet, "author", "A10")
    global_data(data, sheet, "url", "A11")
    global_data(data, sheet, "mission", "A12")
    global_const(data, "category", "geodati")
    pprint.pprint(data)

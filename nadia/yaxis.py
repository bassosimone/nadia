# nadia/yaxis.py

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
# Pass 3.  Guess the format of the Y axis and then append the name of
# the dataset and eventual comments to each package.  This is the second
# pass of making the data a bit more machine friendly, i.e. one-line,
# one-package.
#

import pprint
import sys

if __name__ == "__main__":
    sys.path.insert(0, ".")

from nadia.excel import open_sheet
from nadia.excel import row_values
from nadia.data import data_section

def _is_garbage(x):
    for k in x:
        if k:
            return False
    return True

# garbage-before-data = the garbage is the section of the data
def _y_axis_garbage_before_type(M, data):
    i = 0
    garbage = ""
    data[0].append("Type")
    while i < len(M):
        y_row = M[i][0]
        i = i + 1                               # XXX
        if _is_garbage(data[i]):
            garbage = y_row
            continue
        # do not add "/" if garbage is ""
        value = garbage
        if value:
            value += "/"
        value += y_row
        data[i].append(value)

# type-before-garbage = the garbage is the description of the data
def _y_axis_type_before_garbage(M, data):
    i = 0
    last = -1
    data[0].append("Type")
    data[0].append("Description")
    while i < len(M):
        y_row = M[i][0]
        i = i + 1                               # XXX
        if not _is_garbage(data[i]):
            data[i].append(y_row)
            data[i].append("")
            last = i
            continue
        if last == -1:
            continue
        data[last][-1] += y_row

def y_axis(sheet, range, data):
    M = row_values(sheet, range)
    for R in M:
        if len(R) != 1:
            raise ValueError
    if _is_garbage(data[1]):
        _y_axis_garbage_before_type(M, data)
    else:
        _y_axis_type_before_garbage(M, data)

if __name__ == "__main__":
    sheet = open_sheet("test/sample1.xls")
    data = data_section(sheet, "C3:R6", "C7:R24")
    y_axis(sheet, "B7:B24", data)
    pprint.pprint(data)

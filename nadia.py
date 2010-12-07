#!/usr/bin/env python

#
# Copyright (c) 2010 Simone Basso <bassosimone@gmail.com>
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

import csv, sys
import ConfigParser

try:
    import xlrd, xlwt.Utils
except ImportError:
    import os.path
    where = os.path.dirname(sys.argv[0])
    if not where:
        where = "."
    sys.path.insert(0, where + "/xlrd-0.7.1")
    sys.path.insert(0, where + "/xlwt-0.7.2")
    import xlrd, xlwt.Utils

# wrappers

def open_sheet(filepath):
    workbook = xlrd.open_workbook(filepath)
    sheet = workbook.sheet_by_index(0)
    return sheet

def _convert_range(range):
    v = range.split(":")
    if len(v) != 2:
        raise ValueError
    firstrow, firstcol = xlwt.Utils.cell_to_rowcol2(v[0])
    lastrow, lastcol = xlwt.Utils.cell_to_rowcol2(v[1])
    # sanity
    if (firstrow < 0 or firstcol < 0 or lastrow < 0 or lastcol < 0 or
     firstrow > lastrow or firstcol > lastcol):
        raise ValueError
    return firstrow, firstcol, lastrow, lastcol

def _utf8(elem):
    return unicode(elem).encode("utf-8")

def row_values(sheet, range):
    M = []
    firstrow, firstcol, lastrow, lastcol = _convert_range(range)
    row = firstrow
    while row <= lastrow:
        R = sheet.row_values(row, firstcol, lastcol + 1)
        R = map(_utf8, R)
        M.append(R)
        row = row + 1
    return M

def cell_value(sheet, cell):
    row, col = xlwt.Utils.cell_to_rowcol2(cell)
    value = sheet.cell_value(row, col)
    value = _utf8(value)
    return value

# data section

def _edit_inplace(cells):
    i = 0
    while i < len(cells):
        prev = ""
        j = 0
        while j < len(cells[i]):
            if cells[i][j]:
                prev = cells[i][j]
            #XXX
            if i == 0 or j == 0 or cells[i-1][j-1] == cells[i-1][j]:
                cells[i][j] = prev
            j = j + 1
        i = i + 1

def _compress(cells, separator):
    j = 0
    while j < len(cells[0]):
        i = 1
        while i < len(cells):
            if cells[i][j]:
                cells[0][j] += separator + cells[i][j]
            i = i + 1
        j = j + 1

def data_headers(sheet, range, separator="/"):
    cells = row_values(sheet, range)
    _edit_inplace(cells)
    _compress(cells, separator)
    return cells[0]

data_values = row_values

def data_section(sheet, hdrs_range, values_range):
    hdrs = data_headers(sheet, hdrs_range)
    values = data_values(sheet, values_range)
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

# y axis

def _is_crap(x):
    for k in x:
        if k:
            return False
    return True

# crap-before-data = the crap is the section of the data
def _y_axis_crap_before_type(M, data):
    i = 0
    crap = ""
    data[0].append("Type")
    while i < len(M):
        y_row = M[i][0]
        i = i + 1                               # XXX
        if _is_crap(data[i]):
            crap = y_row
            continue
        data[i].append(crap + "/" + y_row)

# type-before-crap = the crap is the description of the data
def _y_axis_type_before_crap(M, data):
    i = 0
    last = -1
    data[0].append("Type")
    data[0].append("Description")
    while i < len(M):
        y_row = M[i][0]
        i = i + 1                               # XXX
        if not _is_crap(data[i]):
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
    if _is_crap(data[1]):
        _y_axis_crap_before_type(M, data)
    else:
        _y_axis_type_before_crap(M, data)

# empty lines

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

# global data/strings

def global_data(data, sheet, name, cell):
    data[0].append(name.encode("utf-8"))
    value = cell_value(sheet, cell)
    for row in data[1:]:
        row.append(value)

def global_const(data, name, value):
    data[0].append(name.encode("utf-8"))
    for row in data[1:]:
        row.append(value.encode("utf-8"))

# export

def export(data, filename=None):
    if filename:
        fp = open(filename, "wb")
    else:
        fp = sys.stdout
    writer = csv.writer(fp)
    for row in data:
        writer.writerow(row)

# top-level interface

class NadiaProcessor(object):
    def __init__(self, filename, data_headers_location, data_values_location,
     y_axis_location, global_data_vector, global_const_vector):
        sheet = open_sheet(filename)
        data = data_section(sheet, data_headers_location, data_values_location)
        y_axis(sheet, y_axis_location, data)
        data = remove_empty_lines(data)
        for name, location in global_data_vector:
            global_data(data, sheet, name, location)
        for name, value in global_const_vector:
            global_const(data, name, value)
        filename += ".nadia.csv"
        export(data, filename)

class NadiaConfig(ConfigParser.SafeConfigParser):
    def __init__(self, filename):
        ConfigParser.SafeConfigParser.__init__(self)
        self.read([filename])
        filename = self.get("nadia", "filename")
        data_headers_location = self.get("nadia", "data_headers_location")
        data_values_location = self.get("nadia", "data_values_location")
        y_axis_location = self.get("nadia", "y_axis_location")
        global_data_vector = []
        for name, value in self.items("global_data"):
            global_data_vector.append((name, value))
        global_const_vector = []
        for name, value in self.items("global_const"):
            global_const_vector.append((name, value))
        NadiaProcessor(filename, data_headers_location, data_values_location,
         y_axis_location, global_data_vector, global_const_vector)

__all__ = [
    "NadiaProcessor",
    "NadiaConfig",
]

# main

if __name__ == "__main__":
    arguments = sys.argv[1:]
    for arg in arguments:
        NadiaConfig(arg)

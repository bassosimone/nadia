# nadia/main.py

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

import ConfigParser
import sys

if __name__ == "__main__":
    sys.path.insert(0, ".")

from nadia.excel import open_sheet
from nadia.data import data_section
from nadia.yaxis import y_axis
from nadia.finalize import global_data
from nadia.finalize import global_const
from nadia.finalize import remove_empty_lines
from nadia.jsonize import jsonize

def nadia_process_xls(filename, data_headers_location, data_values_location,
  y_axis_location, global_data_vector, global_const_vector, outfile):

    """Given an XLS file and all the settings -- such as where are headers,
       where is data, where is Y axis, etc. -- this function translates such
       XLS file in a sequence of JSON that match very closely the format of
       a CKAN package."""

    sheet = open_sheet(filename)

    data = data_section(sheet, data_headers_location, data_values_location)
    y_axis(sheet, y_axis_location, data)

    data = remove_empty_lines(data)
    for name, location in global_data_vector:
        global_data(data, sheet, name, location)
    for name, value in global_const_vector:
        global_const(data, name, value)

    jsonize(data, outfile)

def nadia_process_cnf(conffile):

    """Given a configuration file that describes the format of an XLS
       nadia file, this function fetches the configuration and invokes
       the XLS processor with the proper parameters."""

    config = ConfigParser.SafeConfigParser()
    config.read([conffile])

    filename = config.get("nadia", "filename")

    data_headers_location = config.get("nadia", "data_headers_location")
    data_values_location = config.get("nadia", "data_values_location")
    y_axis_location = config.get("nadia", "y_axis_location")

    global_data_vector = []
    for name, value in config.items("global_data"):
        global_data_vector.append((name, value))

    global_const_vector = []
    for name, value in config.items("global_const"):
        global_const_vector.append((name, value))

    nadia_process_xls(filename, data_headers_location, data_values_location,
     y_axis_location, global_data_vector, global_const_vector, sys.stdout)

def main(argv):
    arguments = sys.argv[1:]

    if len(arguments) == 0:
        sys.stderr.write("Usage: nadia conf ...\n")
        sys.exit(1)

    for arg in arguments:
        nadia_process_cnf(arg)

if __name__ == "__main__":
    main(sys.argv)

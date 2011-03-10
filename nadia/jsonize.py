# nadia/jsonize.py

#
# Copyright (c) 2011 Simone Basso <bassosimone@gmail.com>
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
# When the dataset is ready, we export each line with JSON.  The export
# tries to group fields in a CKAN-friendly way so that, ideally, one can
# load the json and pass it directly to ckan-client.
#

import sys

try:
    import json
except ImportError:
    try:
        import simplejson as json
    except ImportError:
        sys.stderr.write("nadia: with Python < 2.6 please install simplejson")
        sys.exit(1)

if __name__ == "__main__":
    sys.path.insert(0, ".")

from nadia.excel import open_sheet
from nadia.excel import row_values
from nadia.data import data_section
from nadia.yaxis import y_axis
from nadia.finalize import remove_empty_lines
from nadia.finalize import global_data
from nadia.finalize import global_const

from nadia.django import slugify

class CKANPackage(object):
    def __init__(self):
        self.author = u""
        self.name = u""
        self.url = u""
        self.notes = []         # XXX
        self.tags = []
        self.extras = {}
        self.title = u""

def jsonize(data, fp, indent=None):
    headers = data[0]
    body = data[1:]
    for row in body:
        package = CKANPackage()

        #
        # The algorithm here matches loosely the one that has been
        # implemented in <ckanload-italy-nexa>.
        #

        for j in range(0, len(row)):
            cell = row[j]
            header = headers[j]

            if (header == "datasource" or header == "istituzione" or
              header == "author"):
                package.author = cell
                continue
            if header == "name":
                package.name = cell
                continue
            if header == "url":
                package.url = cell
                continue
            if (header == "tipologia di dati" or
              header == "diritti sul database"):
                package.notes.append(cell)
                continue
            if header == "title":
                package.title = cell
                continue
            if header == "tag":
                package.tags.append(cell.lower().replace(" ", "-"))
                continue

            if cell:
                package.extras[header] = cell

        # XXX
        package.notes = "\n\n".join(package.notes)

        #
        # As suggested by steko, the machine readable name must
        # be prepended with a slugified version of the name of the
        # dataset author.
        # While on that, make sure the author name is not all-
        # uppercase because that looks ugly.
        # Ensure that the package name is not too long or the server
        # will have a boo.
        #

        name = slugify(package.author.lower())
        if not package.name.startswith(name):
            package.name = name + "_" + package.name
        package.name = package.name[:100]

        #
        # AFAIK vars() here will work as long as all the variables of
        # `package` have been initialized using __init__().  This is
        # what the code above already does.  Nonetheless I whish to add
        # this comment for future robusteness of the code.
        #

        octets = json.dumps(vars(package), indent=indent)
        fp.write(octets)
        fp.write("\n")

if __name__ == "__main__":
    sheet = open_sheet("test/sample1.xls")
    data = data_section(sheet, "C3:R6", "C7:R24")
    y_axis(sheet, "B7:B24", data)
    data = remove_empty_lines(data)
    global_data(data, sheet, "author", "A10")
    global_data(data, sheet, "url", "A11")
    global_data(data, sheet, "mission", "A12")
    global_const(data, "category", "geodati")
    jsonize(data, sys.stdout, 4)

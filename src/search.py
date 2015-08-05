#!/usr/bin/python
# encoding: utf-8
#
# Copyright © 2015 deanishe@deanishe.net
#
# MIT Licence. See http://opensource.org/licenses/MIT
#
# Created on 2015-08-05
#

"""
Demo Script Filter script using Excel files as data source.

Reads first two columns of first worksheet in the specified file.
Filters data on contents of first column and outputs the content
of the second column.
"""

from __future__ import print_function, unicode_literals, absolute_import


import sys

from xlrd import open_workbook
from workflow import Workflow, ICON_WARNING, ICON_USER

# Path to the Excel file you want to use as a data source
# Change this to point to the file you want to read
EXCEL_FILE = 'Data.xls'
# EXCEL_FILE = 'Data.xlsx'

log = None


def load_excel_data(filepath):
    """Load first 2 columns from first worksheet of Excel file `filepath`."""
    # Load Excel file at `filepath`
    wb = open_workbook(filepath)
    # Grab first sheet (index 0 = first sheet)
    sheet = wb.sheets()[0]
    log.debug('Sheet name : %s', sheet.name)

    # Read first two columns of first worksheet
    data = []
    for row in range(sheet.nrows):
        # First column (and row) have index 0
        cell1 = sheet.cell(row, 0).value
        cell2 = sheet.cell(row, 1).value
        data.append((cell1, cell2))
        log.debug('Read row : %s, %s', cell1, cell2)

    return data


def main(wf):
    query = None
    if len(wf.args):
        query = wf.args[0]

    # Read data from Excel file
    data = load_excel_data(EXCEL_FILE)

    if query:
        # Filter data on first column
        data = wf.filter(query, data, lambda t: t[0], min_score=30)

    if not data:  # Nothing matches
        wf.add_item('No matching results',
                    'Try a different query',
                    icon=ICON_WARNING)
        wf.send_feedback()

    log.debug('%d result(s) for `%s`', len(data), query)

    # Send results to Alfred
    for name, id_ in data:
        sub = '↩ to copy ID to clipboard / ⌘+↩ to paste ID to active app'
        wf.add_item('{0} // {1}'.format(name, id_),
                    sub,
                    arg=id_,
                    valid=True,
                    icon=ICON_USER,
                    largetext=id_,
                    copytext=id_)

    wf.send_feedback()


if __name__ == '__main__':
    wf = Workflow()
    log = wf.logger
    sys.exit(wf.run(main))

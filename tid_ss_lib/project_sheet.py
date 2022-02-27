#-----------------------------------------------------------------------------
# Title      : Manipulate Project Sheet
#-----------------------------------------------------------------------------
# This file is part of the TID ID Smartsheets software platform. It is subject to
# the license terms in the LICENSE.txt file found in the top-level directory
# of this distribution and at:
#    https://confluence.slac.stanford.edu/display/ppareg/LICENSE.html.
# No part of the TID ID Smartsheets software platform, including this file, may be
# copied, modified, propagated, or distributed except according to the terms
# contained in the LICENSE.txt file.
#-----------------------------------------------------------------------------
import smartsheet
from . import navigate

# Set formats
#
# https://smartsheet-platform.github.io/api-docs/#formatting
#
# Colors 31 = Dark Blue
#        26 = Dark Gray
#        23 = Blue
#        18 - Gray
#           = White


def get_project_list(*, client):

    sheet = client.Sheets.get_sheet(navigate.TID_ID_LIST_SHEET, include='format')

    ret = []

    for row in sheet.rows:

        if row.cells[0].value is not None and row.cells[0].value != '' and \
           row.cells[1].value is not None and row.cells[1].value != '' and \
           row.cells[2].value is not None and row.cells[2].value != '' and \
           row.cells[3].value is not None and row.cells[3].value != '' and \
           row.cells[5].value is not None and row.cells[5].value != '':

            proj = {'program': row.cells[0].value,
                    'name': row.cells[1].value,
                    'pa_number': row.cells[2].value,
                    'id': int(row.cells[3].value),
                    'pm': row.cells[5].value}

            ret.append(proj)

    return ret

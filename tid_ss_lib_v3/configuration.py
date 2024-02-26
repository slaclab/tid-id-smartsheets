#-----------------------------------------------------------------------------
# Title      : Top Level Configuration File
#-----------------------------------------------------------------------------
# This file is part of the TID ID Smartsheets software platform. It is subject to
# the license terms in the LICENSE.txt file found in the top-level directory
# of this distribution and at:
#    https://confluence.slac.stanford.edu/display/ppareg/LICENSE.html.
# No part of the TID ID Smartsheets software platform, including this file, may be
# copied, modified, propagated, or distributed except according to the terms
# contained in the LICENSE.txt file.
#-----------------------------------------------------------------------------

from types import SimpleNamespace
import smartsheet

division_list = {'id' : 1446826922168196 }

# Get division configuration information
def get_division(*, client, div):

    if div not in division_list:
        raise Exception("Invalid division")

    sheet = client.Sheets.get_sheet(division_list[div])

    config = {'configuration' : division_list['id'],
              'wbs_exports'   : {},
              'ms_exports'    : {},
              'key'           : div }

    # Process the rows
    for rowIdx in range(0,len(sheet.rows)):
        key = sheet.rows[rowIdx].cells[0].value

        if key is not None and key != "":

            if key == 'Earned Value Date':
                config['earned_value_date'] = sheet.rows[rowIdx].cells[2].value

            elif key.startswith('WBS Export'):
                config['wbs_exports'][key.split(' ')[2]] = sheet.rows[rowIdx].cells[1].value

            elif key.startswith('MS Export'):
                config['ms_exports'][key.split(' ')[2]] = sheet.rows[rowIdx].cells[1].value

            else:
                k = key.replace(' ', '_').lower()
                config[k] = sheet.rows[rowIdx].cells[1].value

    return SimpleNamespace(**config)


def get_user_map(*, client, div):
    userMap = {}

    sheet = client.Sheets.get_sheet(int(div.wbs_name_lookup))

    # Process the rows
    for rowIdx in range(0,len(sheet.rows)):
        key = sheet.rows[rowIdx].cells[0].value
        value = sheet.rows[rowIdx].cells[1].value

        userMap[key] = value

    return userMap


def add_missing_user_map(*, client, div, miss):
    sheet = client.Sheets.get_sheet(int(div.wbs_name_lookup))
    addRows = []

    for k in miss:

        new_row = smartsheet.models.Row()
        new_row.to_bottom = True

        new_cell = smartsheet.models.Cell()
        new_cell.column_id = sheet.columns[0].id
        new_cell.value = k
        new_cell.strict = False
        new_row.cells.append(new_cell)

        new_cell = smartsheet.models.Cell()
        new_cell.column_id = sheet.columns[1].id
        new_cell.value = k
        new_cell.strict = False
        new_row.cells.append(new_cell)

        addRows.append(new_row)

    client.Sheets.add_rows(sheet.id, addRows)


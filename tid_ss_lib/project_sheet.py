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
                    'id': int(row.cells[1].value),
                    'name': row.cells[2].value,
                    'pa_number': row.cells[3].value,
                    'pm': row.cells[5].value}

            ret.append(proj)

    return ret


def check_cell_value(*, client, sheet, rowIdx, row, col, expect):
    new_cell = None

    if row.cells[col].value != expect:
        print(f"    Row {rowIdx+1} Cell {col+1} value mismatch. Got {row.cells[col].value} expect {expect}")
        new_cell = smartsheet.models.Cell()
        new_cell.column_id = sheet.columns[col].id
        new_cell.value = expect
        new_cell.strict = False

    return new_cell


def check_cell_formula(*, client, sheet, rowIdx, row, col, expect):
    new_cell = None

    if row.cells[col].formula != expect:
        print(f"    Row {rowIdx+1} Cell {col+1} formula mismatch. Got {row.cells[col].formula} expect {expect}")
        new_cell = smartsheet.models.Cell()
        new_cell.column_id = sheet.columns[col].id
        new_cell.formula = expect
        new_cell.strict = False

    return new_cell


def check_row(*, client, sheet, rowIdx, folderList, doFixes):

    row = sheet.rows[rowIdx]

    new_row = smartsheet.models.Row()
    new_row.id = row.id

    # First we get the project ID
    fid = int(row.cells[1].value)

    if fid not in folderList:
        print(f"    Row {rowIdx+1} contains unknown project ID {fid}")
        return

    p = folderList[fid]
    p['tracked'] =True

    # Project Name
    ret = check_cell_value(client=client, sheet=sheet, rowIdx=rowIdx, row=row, col=2, expect=p['name'])

    if ret is not None:
        new_row.cells.append(ret)

    # Status Month
    if rowIdx != 0:
        ret = check_cell_formula(client=client, sheet=sheet, rowIdx=rowIdx, row=row, col=8, expect='=[Status Month]1')

        if ret is not None:
            new_row.cells.append(ret)

    # Lookup Fields
    for col in range(9, 24):
        exp = "=VLOOKUP([Status Month]@row, {"
        exp += p['name']
        exp += " Tracking Range 1}, "
        exp += str(3 + (col-9))
        exp += ", false)"

        ret = check_cell_formula(client=client, sheet=sheet, rowIdx=rowIdx, row=row, col=col, expect=exp)

        if ret is not None:
            new_row.cells.append(ret)

    # Check hyperlink Column
    col = 24
    if row.cells[col].hyperlink is None or row.cells[col].hyperlink.url != p['url'] or row.cells[col].value != p['path']:

        if row.cells[col].hyperlink is None:
            print(f"    Row {rowIdx+1} cell {col+1} missing hyperlink")

        elif row.cells[col].hyperlink.url != p['url']:
            print(f"    Row {rowIdx+1} cell {col+1} hyperlink url mismatch. Got {row.cells[col].hyperlink.url} expect {p['url']}")

        elif row.cells[col].value != p['path']:
            print(f"    Row {rowIdx+1} cell {col+1} hyperlink value mismatch. Got {row.cells[col].value} expect {p['path']}")

        new_cell = smartsheet.models.Cell()
        new_cell.column_id = sheet.columns[col].id
        new_cell.value = p['path']
        new_cell.hyperlink = smartsheet.models.Hyperlink()
        new_cell.hyperlink.url = p['url']
        new_cell.strict = False
        new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check(*, client, maxRow, doFixes):
    sheet = client.Sheets.get_sheet(navigate.TID_ID_LIST_SHEET, include='format')

    # Get folder list:
    folderList = navigate.get_active_list(client=client)

    for row in range(0, maxRow):
        check_row(client=client, sheet=sheet, rowIdx=row, folderList=folderList, doFixes=doFixes)

    for k, v in folderList.items():
        if v['tracked'] is False:
            print(f"    Project {v['name']} with id {k} is not tracked")

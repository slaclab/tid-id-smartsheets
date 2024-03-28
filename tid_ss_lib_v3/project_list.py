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

def get_project_list(*, client, div):

    sheet = client.Sheets.get_sheet(int(div.project_list), include='format')

    ret = []

    for row in sheet.rows:
        proj = {}

        proj['name']      = row.cells[0].value  if row.cells[0].value  is not None else ''
        proj['program']   = row.cells[1].value  if row.cells[1].value  is not None else 'Unkown'
        proj['pm']        = row.cells[3].value  if row.cells[3].value  is not None else 'Unknown'
        proj['id']        = int(row.cells[6].value) if row.cells[6].value  is not None else ''
        proj['updated']   = row.cells[23].value if row.cells[23].value is not None else ''

        if proj['name'] != '' and proj['id'] != '' and proj['updated'] == 'Yes':
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
    fid = int(row.cells[7].value)

    if fid not in folderList:
        print(f"    Row {rowIdx+1} contains unknown project ID {fid}")
        return

    p = folderList[fid]
    p['tracked'] = True

    # Project Name
    ret = check_cell_value(client=client, sheet=sheet, rowIdx=rowIdx, row=row, col=1, expect=p['name'])

    if ret is not None:
        new_row.cells.append(ret)

    if len(p['name']) > 30:
        print(f"    Project name {p['name']} is too long")

    # Status Month
    if rowIdx != 0:
        ret = check_cell_formula(client=client, sheet=sheet, rowIdx=rowIdx, row=row, col=8, expect='=[Status Month]1')

        if ret is not None:
            new_row.cells.append(ret)

    LookupIndexes = { 8: 3,  # Total Budget
                      9: 2,  # Actual Cost
                     10: 4,  # Remaining Funds
                     11: 7,  # Cost Variance
                     12: 8,  # CPI
                     13: 9,  # Schedule Variance
                     14: 10, # SPI
                     15: 11, # Budget Risk
                     16: 12, # Schedule Risk
                     17: 13, # Scope    Risk
                     18: 14} # Description Of Status

    for col, enum in LookupIndexes.items():
        exp = "=VLOOKUP([Status Month]@row, {"
        exp += p['name']
        exp += " Tracking Range 1}, "
        exp += str(enum)
        exp += ", false)"

        ret = check_cell_formula(client=client, sheet=sheet, rowIdx=rowIdx, row=row, col=col, expect=exp)

        if ret is not None:
            new_row.cells.append(ret)

    # Check hyperlink Column
    col = 20
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

    # Check Budget Index
    col = 22
    exp = '=[Total Budget]@row / [Real Budget]@row'

    ret = check_cell_formula(client=client, sheet=sheet, rowIdx=rowIdx, row=row, col=col, expect=exp)

    if ret is not None:
        new_row.cells.append(ret)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check(*, client, doFixes, div):


    # Get folder list:
    print("Searching active directory for projects ...")
    #folderList = navigate.get_active_list(client=client,div=div)

    print("Processing division project sheet ...\n")
    #sheet = client.Sheets.get_sheet(div.project_list, include='format')
    sheet = client.Sheets.get_sheet(div.project_list)

    for rowIdx in range(len(sheet.rows)):

        # Skip rows that don't have a project ID
        if sheet.rows[rowIdx].cells[7].value is not None and sheet.rows[rowIdx].cells[7].value != "":
            check_row(client=client, sheet=sheet, rowIdx=rowIdx, folderList=folderList, doFixes=doFixes)

    for k, v in folderList.items():
        if v['tracked'] is False:
            print(f"    Project {v['name']} with id {k} is not tracked")



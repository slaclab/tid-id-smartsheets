#-----------------------------------------------------------------------------
# Title      : Actuals Sheet Manipulation
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

import datetime

def check(*, client, sheet, doFixes, resources):
    inLabor = False

    resFound = {v : False for v in resources}
    empty = []
    addRows = []
    parentId = None

    # Process the rows
    for rowIdx in range(0,len(sheet.rows)):

        if sheet.rows[rowIdx].cells[0].value == 'Labor Actuals':
            inLabor = True
            parentId = sheet.rows[rowIdx].id

        elif inLabor:
            resName = sheet.rows[rowIdx].cells[1].value

            if isinstance(resName, str) and resName != "":
                resFound[resName] = True
            else:
                empty.append(rowIdx)

    # Add new resource rows
    for k,v in resFound.items():

        if v is False:
            if doFixes:
                print(f"   Adding missing resource to actuals: {k}")
                new_row = smartsheet.models.Row()
                new_row.parent_id = parentId

                # Item
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[0].id
                new_cell.value = k
                new_cell.strict = False
                new_row.cells.append(new_cell)

                # Assigned To
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[1].id
                new_cell.value = k
                new_cell.strict = False
                new_row.cells.append(new_cell)

                # Actual Hours
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[2].id
                new_cell.value = "0"
                new_cell.strict = False
                new_row.cells.append(new_cell)

                # Actual Cost
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[3].id
                new_cell.value = "0"
                new_cell.strict = False
                new_row.cells.append(new_cell)
                addRows.append(new_row)

            else:
                print(f"   Missing resource in actuals: {k}")

    # Process add rows
    if len(addRows) > 0:
        client.Sheets.add_rows(sheet.id, addRows)


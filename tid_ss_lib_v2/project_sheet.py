#-----------------------------------------------------------------------------
# Title      : Schedule Sheet Manipulation
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
from . import project_sheet_columns


def find_columns(*, client, sheet, doFixes, cData):

    for k,v in cData.items():
        found = False
        for i in range(len(sheet.columns)):
            if sheet.columns[i].title == k:
                found = True
                if v['position'] != i:
                    print(f"Column location mismatch for {k}. Expected at {v['position']}, found at {i}.")
                    v['position'] = i

                #print(f"Found column {k} at position {v['position']}, with id {sheet.columns[i].id}")

        if not found:
            print(f"Column not found: {k}.")

            if doFixes is True:
                print(f"Adding column: {k}.")
                col = smartsheet.models.Column({'title': k,
                                                'type': v['type'],
                                                'index': v['position']})

                client.Sheets.add_columns(sheet.id, [col])

            return False

    return True

def check_row(*, client, sheet, rowIdx, key, div, cData, doFixes):

    if div == 'id':
        laborRate = navigate.TID_ID_RATE_NOTE
    elif div == 'cds':
        laborRate = navigate.TID_CDS_RATE_NOTE

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for k,v in cData.items():
        doFormat = False
        idx = v['position']
        data = v[key]

        # First check format
        if data['format'] is not None and row.cells[idx].format != data['format']:
            print(f"   Incorrect format in row {rowIdx+1} {key} cell {idx+1} in project file. Got {row.cells[idx].format} Exp {data['format']}")
            doFormat = True

        # Row has a formula
        if data['formula'] is not None:
            doRow = False

            if not hasattr(row.cells[idx],'formula') or row.cells[idx].formula != data['formula']:
                print(f"   Incorrect formula in row {rowIdx+1} {key} cell {idx+1} in project file. Got {row.cells[idx].formula} Exp {data['formula']}")
                doRow = True

            if doRow or doFormat:
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[idx].id
                new_cell.formula = data['formula']

                if data['format'] is not None:
                    new_cell.format = data['format']

                new_cell.strict = False
                new_row.cells.append(new_cell)

        # Row has a forced value
        elif data['forced'] is not None:
            doRow = False

            if data['forced'] == 'MS_OVERHEAD':
                data['forced'] = navigate.OVERHEAD_NOTE
            elif data['forced'] == 'LAB_RATE':
                data['forced'] = laborRate

            if (not (row.cells[idx].value is None and data['forced'] == '')) and (row.cells[idx].value is None or row.cells[idx].value != data['forced']):
                print(f"   Incorrect value in row {rowIdx+1} {key} cell {idx+1} in project file. Got {row.cells[idx].value} Exp {data['forced']}")
                doRow = True

            if doRow or doFormat:
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[idx].id
                new_cell.value = data['forced']

                if data['format'] is not None:
                    new_cell.format = data['format']

                new_cell.strict = False
                new_row.cells.append(new_cell)

        # Row has a default value and is empty
        elif data['default'] is not None and (row.cells[idx].value is None or row.cells[idx].value == ''):
            print(f"   Missing default value in row {rowIdx+1} {key} cell {idx+1} in project file.")

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[idx].id
            new_cell.value = data['default']

            if data['format'] is not None:
                new_cell.format = data['format']

            new_cell.strict = False
            new_row.cells.append(new_cell)

        # Catch format only case
        elif doFormat:

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[idx].id

            if row.cells[idx].value is None:
                new_cell.value = ''
            else:
                new_cell.value = row.cells[idx].value
            new_cell.format = data['format']
            new_cell.strict = False
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to project row {rowIdx+1} {key}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check(*, client, sheet, doFixes, div):
    inLabor = False
    inMS = False
    cData = project_sheet_columns.ColData

    # First check structure
    while True:
        ret = find_columns(client=client, sheet=sheet, doFixes=doFixes, cData=cData)

        if ret is False and doFixes is False:
            return False

        elif ret is True:
            break

        else:
            sheet = client.Sheets.get_sheet(sheet.id, include='format')

    # First do row 0
    check_row(client=client, sheet=sheet, rowIdx=0, key='top', div=div, cData=cData, doFixes=doFixes)

    # First walk through the rows and create a parent list
    parents = set()
    for rowIdx in range(len(sheet.rows)):
        parents.add(sheet.rows[rowIdx].parent_id)

    # Process the rest of the rows
    for rowIdx in range(1,len(sheet.rows)):

        # MS Section
        if sheet.rows[rowIdx].cells[0].value == 'M&S':
            check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='ms_top', div=div, cData=cData, doFixes=doFixes)
            inMS = True
            inLabor = False
        elif sheet.rows[rowIdx].cells[0].value == 'Labor':
            check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='labor_top', div=div, cData=cData, doFixes=doFixes)
            inMS = False
            inLabor = True
        elif inMS or inLabor:

            if inMS:
                key = 'ms'
            else:
                key = 'labor'

            if sheet.rows[rowIdx].id in parents:
                key += '_parent'
            else:
                key += '_task'

            check_row(client=client, sheet=sheet, rowIdx=rowIdx, key=key, div=div, cData=cData, doFixes=doFixes)

    return True

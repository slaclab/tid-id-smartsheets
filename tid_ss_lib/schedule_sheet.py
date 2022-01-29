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

# TODO:
#    Color format parent rows
#    Fix indent issue


def check_structure(*, sheet):

    columns = ['Task Name From Budget',
               'Predecessors',
               'Start',
               'End',
               'Duration',
               'Baseline Start',
               'Baseline Finish',
               'Baseline Duration (days)',
               'Planned Labor Hours From Budget',
               'Assigned To',
               '% Effort Planned From Resource',
               'Baseline Variance',
               'Duration Variance',
               'Actual Labor Hours',
               '% Complete',
               'Notes',
               'PA Number']

    # Check column count
    if len(sheet.columns) != len(columns):
        print(f"   Wrong number of columns in schedule file: Got {len(sheet.columns)}.")
        return False

    else:
        ret = True

        for i,v in enumerate(columns):
            if sheet.columns[i].title != v:
                print(f"   Mismatch schedule column name for col {i+1}. Got {sheet.columns[i].title}. Expect {v}.")
                ret = False

        return ret




def check_parent_row(*, client, sheet, rowIdx, doFixes):
    formulas = { 7: '=NETWORKDAYS([Baseline Start]@row, [Baseline Finish]@row)',
                12: '=SUM(CHILDREN())',
                13: '=SUM(CHILDREN())' }

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for i in range(1,len(row.cells)):
        if i in formulas:
            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i]:
                print(f"   Invalid value in row {rowIdx+1} col {i+1} in schedule file. Expected '{formulas[i]}'. Got '{row.cells[i].formula}'.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = formulas[i]
                new_cell.strict = False
                new_row.cells.append(new_cell)

        elif row.cells[i].link_in_from_cell is not None:
            print(f"   Invalid sheet link in row {rowIdx+1} cell {i+1} in schedule file.")
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[i].id
            new_cell.value = ''
            new_cell.strict = False
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_row(*, client, sheet, rowIdx, doFixes):
    formulas = {  7: '=NETWORKDAYS([Baseline Start]@row, [Baseline Finish]@row)',
                 10: '=([Planned Labor Hours From Budget]@row / 8) / [Baseline Duration (days)]@row',
                 12: '=Duration@row - [Baseline Duration (days)]@row', }

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for i in range(len(row.cells)):
        if i in formulas:

            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i]:
                print(f"   Invalid value in row {rowIdx+1} col {i+1} in schedule file. Expected '{formulas[i]}'. Got '{row.cells[i].formula}'.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = formulas[i]
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_links(*, client, sheet, rowIdx, laborRows, laborSheet, doFixes):
    links = { 8: 2, 16: 20 }

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for i in range(1,len(row.cells)):
        if i in links:
            rowIdTar = laborRows[rowIdx-1]['data'].id
            colIdTar = laborRows[rowIdx-1]['data'].cells[links[i]].column_id
            shtIdTar = laborSheet.id

            if row.cells[i].link_in_from_cell is None or \
                row.cells[i].link_in_from_cell.row_id != rowIdTar or \
                row.cells[i].link_in_from_cell.column_id != colIdTar or \
                row.cells[i].link_in_from_cell.sheet_id != shtIdTar:

                print(f"   Incorrect schedule link for row {rowIdx+1} column {i+1}.")

                cell_link = smartsheet.models.CellLink()
                cell_link.sheet_id = shtIdTar
                cell_link.row_id = rowIdTar
                cell_link.column_id = colIdTar

                new_cell = smartsheet.models.Cell()
                new_cell.column_id = row.cells[i].column_id
                new_cell.value = smartsheet.models.ExplicitNull()
                new_cell.link_in_from_cell = cell_link
                new_row.cells.append(new_cell)

        elif row.cells[i].link_in_from_cell is not None:
            print(f"   Invalid schedule link for row {rowIdx+1} column {i+1}.")
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = row.cells[i].column_id
            new_cell.value = ''
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


# Find stale or broken links to budget file
def check_broken(*, client, sheet, doFixes):
    delList = []

    for rowIdx in range(1,len(sheet.rows)):
        row = sheet.rows[rowIdx]

        if row.cells[0].link_in_from_cell is None or row.cells[0].link_in_from_cell.status != 'OK':
            print(f"   Found stale row {rowIdx+1} in schedule file.")
            delList.append(row.id)

    if doFixes and len(delList) != 0:
        print(f"   Deleting stale rows.")
        pass
        client.Sheets.delete_rows(sheet.id, delList)


# Check matching rows to budget file
def check_rows(*, client, sheet, laborRows, laborSheet, doFixes):
    lastId = sheet.rows[0].id
    rowIdx = 1

    for i in range(len(laborRows)):
        if rowIdx < len(sheet.rows) and \
             sheet.rows[rowIdx].cells[0].link_in_from_cell is not None and \
             sheet.rows[rowIdx].cells[0].link_in_from_cell.row_id == laborRows[i]['data'].id:

            lastId = sheet.rows[rowIdx].id
            laborRows[i]['link'] = sheet.rows[rowIdx]
            rowIdx += 1
        else:
            print(f"   Missing schedule row linked to budget row = {laborRows[i]['data'].row_number}, parent = {laborRows[i]['parent']}")

            if not doFixes:
                print("   Skipping further row checks.")
                return False

            print("   Adding new row at position {rowIdx+1}.")

            newRow = smartsheet.models.Row()

            if rowIdx == 1:
                newRow.parent_id = lastId
            else:
                newRow.sibling_id = lastId

            newRow.cells.append({'column_id': sheet.rows[0].cells[0].column_id, 'value': ''})

            resp = client.Sheets.add_rows(sheet.id, [newRow])
            lastId = resp.data[0].id

            new_row = smartsheet.models.Row()
            new_row.id = lastId

            cell_link = smartsheet.models.CellLink()
            cell_link.sheet_id = laborSheet.id
            cell_link.row_id = laborRows[i]['data'].id
            cell_link.column_id = laborRows[i]['data'].cells[0].column_id

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.rows[0].cells[0].column_id
            new_cell.value = smartsheet.models.ExplicitNull()
            new_cell.link_in_from_cell = cell_link
            new_row.cells.append(new_cell)
            resp = client.Sheets.update_rows(sheet.id, [new_row])
            return True

    return False


def check(*, client, sheet, laborRows, laborSheet, doFixes):

    check_broken(client=client, sheet=sheet, doFixes=doFixes)
    sheet = client.Sheets.get_sheet(sheet.id)

    while check_rows(client=client, sheet=sheet, laborRows=laborRows, laborSheet=laborSheet, doFixes=doFixes):
        sheet = client.Sheets.get_sheet(sheet.id)

    # Process top row
    check_parent_row(client=client, sheet=sheet, rowIdx=0, doFixes=doFixes)

    # Process rows
    for i in range(len(laborRows)):

        if laborRows[i]['parent']:
            check_parent_row(client=client, sheet=sheet, rowIdx=i+1, doFixes=doFixes)
        else:
            check_task_row(client=client, sheet=sheet, rowIdx=i+1, doFixes=doFixes)
            check_task_links(client=client, sheet=sheet, rowIdx=i+1, laborRows=laborRows, laborSheet=laborSheet, doFixes=doFixes)

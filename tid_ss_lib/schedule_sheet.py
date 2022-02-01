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

formGray  = ",,,,,,,,,18,,,,,,,"
formDGray = ",,,,,,,,,26,,,,,,,"
formWhite = ",,,,,,,,,,,,,,,1,"
formBlue  = ",,,,,,,,,23,,,,,,,"
formDBlue = ",,1,,,,,,2,39,,,,,,1,"
formGreen = ",,,,,,,,,22,,,,,,,"

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


def check_parent_row(*, client, sheet, rowIdx, doFixes, title):
    formulas = { 7: '=NETWORKDAYS([Baseline Start]@row, [Baseline Finish]@row)',
                12: '=SUM(CHILDREN())',
                13: '=SUM(CHILDREN())' }

    # These Columns SHould Have No Value
    noValue = set([8, 9, 14, 16])

    # Preserve Values
    noChange = set([1])

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    # Check first Cell of first row
    if title is not None:
        form = formDBlue

        if row.cells[0].value != title or row.cells[0].format != form:
            print(f"   Incorrect row {rowIdx+1}, col 1 value or shading in schedule file. Got '{row.cells[0].value}'. Expected '{title}'.")
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[0].id
            new_cell.value = title
            new_cell.format = form
            new_cell.strict = False
            new_row.cells.append(new_cell)

    else:
        form = formBlue

    new_row.format = form

    for i in range(1,len(row.cells)):
        if i in formulas:

            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i] or row.cells[i].format != form:
                print(f"   Invalid value or shading in row {rowIdx+1} col {i+1} in schedule file. Expected '{formulas[i]}'. Got '{row.cells[i].formula}'.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = formulas[i]
                new_cell.format = form
                new_cell.strict = False
                new_row.cells.append(new_cell)

        if i in noValue:
            if row.cells[i].link_in_from_cell is not None or row.cells[i].format != form or (row.cells[i].value is not None and row.cells[i].value != ''):
                print(f"   Invalid sheet link, shading or value in row {rowIdx+1} cell {i+1} in schedule file.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.value = ''
                new_cell.format = form
                new_cell.strict = False
                new_row.cells.append(new_cell)

        if i in noChange:
            if row.cells[i].format != form:
                print(f"   Invalid sheet shading in row {rowIdx+1} cell {i+1} in schedule file.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id

                if row.cells[i].value is None:
                    new_cell.value = ''
                else:
                    new_cell.value = row.cells[i].value
                new_cell.format = form
                new_cell.strict = False
                new_row.cells.append(new_cell)


    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to schedule row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_row(*, client, sheet, rowIdx, doFixes):
    formulas = {  7: '=NETWORKDAYS([Baseline Start]@row, [Baseline Finish]@row)',
                 10: '=([Planned Labor Hours From Budget]@row / 8) / [Baseline Duration (days)]@row',
                 12: '=Duration@row - [Baseline Duration (days)]@row', }

    init = { 13 : '0',
             14 : '0' }

    status = set([ 15 ])

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for i in range(len(row.cells)):
        if i in formulas:

            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i] or row.cells[i].format != formGray:
                print(f"   Invalid value or shading in row {rowIdx+1} col {i+1} in schedule file. Expected '{formulas[i]}'. Got '{row.cells[i].formula}'.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = formulas[i]
                new_cell.strict = False
                new_cell.format = formGray
                new_row.cells.append(new_cell)

        if i in init:
            if row.cells[i].value is None or row.cells[i].value == '' or row.cells[i].format != formGreen:
                print(f"   Missing default value or shading in row {rowIdx+1} col {i+1} in schedule file.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.format = formGreen
                new_cell.strict = False

                if row.cells[i].value is not None and row.cells[i].value != '':
                    new_cell.value = row.cells[i].value
                else:
                    new_cell.value = init[i]

                new_row.cells.append(new_cell)

        if i in status:
            if row.cells[i].format != formGreen:
                print(f"   Missing shading in row {rowIdx+1} col {i+1} in schedule file.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id

                if row.cells[i].value is not None and row.cells[i].value != '':
                    new_cell.value = row.cells[i].value
                else:
                    new_cell.value = ''

                new_cell.format = formGreen
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_links(*, client, sheet, rowIdx, laborRows, laborSheet, doFixes):
    links = { 8: 2, 16: 20 }

    row = sheet.rows[rowIdx]

    # Need re-link
    relink = set()

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    # First check shading
    for i in range(1,len(row.cells)):
        if i in links:

            if row.cells[i].format != formGray:
                print(f"   Incorrect shading for row {row.row_number} column {i+1}.")
                relink.add(i)
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = row.cells[i].column_id
                new_cell.value = smartsheet.models.ExplicitNull()
                new_cell.format = formGray
                new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to schedule row {row.row_number}.")
        client.Sheets.update_rows(sheet.id, [new_row])

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for i in range(1,len(row.cells)):
        if i in links:
            rowIdTar = laborRows[rowIdx-1]['data'].id
            colIdTar = laborRows[rowIdx-1]['data'].cells[links[i]].column_id
            shtIdTar = laborSheet.id

            if i in relink or row.cells[i].link_in_from_cell is None or \
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
            new_cell.format = formWhite
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

            print(f"   Adding new row at position {rowIdx+1}.")

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
    sheet = client.Sheets.get_sheet(sheet.id, include='format')

    while check_rows(client=client, sheet=sheet, laborRows=laborRows, laborSheet=laborSheet, doFixes=doFixes):
        sheet = client.Sheets.get_sheet(sheet.id, include='format')

    # Process top row
    check_parent_row(client=client, sheet=sheet, rowIdx=0, doFixes=doFixes, title='Labor Task')

    # Process rows
    for i in range(len(laborRows)):

        if laborRows[i]['parent']:
            check_parent_row(client=client, sheet=sheet, rowIdx=i+1, doFixes=doFixes, title=None)
        else:
            check_task_row(client=client, sheet=sheet, rowIdx=i+1, doFixes=doFixes)
            check_task_links(client=client, sheet=sheet, rowIdx=i+1, laborRows=laborRows, laborSheet=laborSheet, doFixes=doFixes)

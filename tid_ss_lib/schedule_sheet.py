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

# Set formats for columns looking at major, minor and task rows
#
# https://smartsheet-platform.github.io/api-docs/#formatting
#
# Colors 39 = Dark Blue
#        26 = Dark Gray
#        23 = Blue
#        22 = Green
#        18 - Gray
#           = White

formatMajor = [",,1,,,,,,2,39,,,,,,,",  # 0
               ",,1,,,,,,2,39,,,,,,,",  # 1
               None,  # 2
               None,  # 3
               None,  # 4
               None,  # 5
               None,  # 6
               ",,1,,,,,,2,39,,,,,,,",  # 7
               ",,1,,,,,,2,39,,,,,,,",  # 8
               ",,1,,,,,,2,39,,,,,,,",  # 9
               None,  # 10
               None,  # 11
               ",,1,,,,,,2,39,,,0,,,,",  # 12
               ",,1,,,,,,2,39,,,0,,,,",  # 13
               ",,1,,,,,,2,39,,,,,3,,",  # 14
               ",,1,,,,,,2,39,,,,,,,",   # 15
               ",,1,,,,,,2,39,,,,,,,",   # 16
               ",,1,,,,,,2,39,,,,,,,",   # 17
               ",,1,,,,,,2,39,,,,,,,",   # 18
               ",,1,,,,,,2,39,,,,,,,",   # 19
               ",,1,,,,,,2,39,,,,,,,"]   # 20

formatMinor = [",,,,,,,,,23,,,,,,,",  # 0
               ",,,,,,,,,23,,,,,,,",  # 1
               None,  # 2
               None,  # 3
               None,  # 4
               None,  # 5
               None,  # 6
               ",,,,,,,,,23,,,,,,,",  # 7
               ",,,,,,,,,23,,,,,,,",  # 8
               ",,,,,,,,,23,,,,,,,",  # 9
               None,  # 10
               None,  # 11
               ",,,,,,,,,23,,,0,,,,",  # 12
               ",,,,,,,,,23,,,0,,,,",  # 13
               ",,,,,,,,,23,,,,,3,,",  # 14
               ",,,,,,,,,23,,,,,,,",   # 15
               ",,,,,,,,,23,,,,,,,",   # 16
               ",,,,,,,,,23,,,,,,,",   # 17
               ",,,,,,,,,23,,,,,,,",   # 18
               ",,,,,,,,,23,,,,,,,",   # 19
               ",,,,,,,,,23,,,,,,,"]   # 20

formatTask = [",,,,,,,,,2,,,,,,,",  # 0
              ",,,,,,,,,2,,,,,,,",  # 1
              None,  # 2
              None,  # 3
              None,  # 4
              None,  # 5
              None,  # 6
              ",,,,,,,,,18,,,,,,,",   # 7
              ",,,,,,,,,18,,,,,,,",   # 8
              ",,,,,,,,,2,,,,,,,",    # 9
              ",,,,,,,,,18,,,,,3,,",  # 10
              None,  # 11
              ",,,,,,,,,18,,,0,,,,",  # 12
              ",,,,,,,,,2,,,0,,,,",  # 13
              ",,,,,,,,,2,,,,,3,,",  # 14
              ",,,,,,,,,22,,,0,,,,",  # 15
              ",,,,,,,,,22,,,,3,,,",  # 16
              ",,,,,,,,,22,,,,,,,",   # 17
              ",,,,,,,,,22,,,,,,,",   # 18
              ",,,,,,,,,22,,,,,,1,",  # 19
              ",,,,,,,,,18,,,,,,,"]   # 20

# Due to a limitation in the API the following columns can't be reformatted through the API:
# Column 1, 3, 4, 5, 6, 7, 11, 12

def fix_structure(*, client, sheet):

    if len(sheet.columns) != 17:
        print(f"   Wrong number of columns in schedule file, could not fix: Got {len(sheet.columns)}.")
        return False

    # Add new columns
    col15 = smartsheet.models.Column({'title': 'Engineer Labor Hours',
                                      'type': 'TEXT_NUMBER',
                                      'index': 15})

    col16 = smartsheet.models.Column({'title': 'Engineer % Complete',
                                      'type': 'TEXT_NUMBER',
                                      'index': 15})

    col17 = smartsheet.models.Column({'title': 'Engineer Finish Date',
                                      'type': 'DATE',
                                      'index': 15})

    col18 = smartsheet.models.Column({'title': 'Engineer Status Ready',
                                      'type': 'PICKLIST',
                                      'options': ['No', 'Yes'],
                                      'index': 15})

    client.Sheets.add_columns(sheet.id, [col15, col16, col17, col18])
    return True

def check_structure(*, sheet):

    columns = ['Task Name From Budget',           # 0
               'Predecessors',                    # 1
               'Start',                           # 2
               'End',                             # 3
               'Duration',                        # 4
               'Baseline Start',                  # 5
               'Baseline Finish',                 # 6
               'Baseline Duration (days)',        # 7
               'Planned Labor Hours From Budget', # 8
               'Assigned To',                     # 9
               '% Effort Planned From Resource',  # 10
               'Baseline Variance',               # 11
               'Duration Variance',               # 12
               'Actual Labor Hours',              # 13
               '% Complete',                      # 14
               'Engineer Labor Hours',            # 15
               'Engineer % Complete',             # 16
               'Engineer Finish Date',            # 17
               'Engineer Status Ready',           # 18
               'Notes',                           # 19
               'PA Number']                       # 20

    # Check column count
    if len(sheet.columns) != len(columns):
        print(f"   Wrong number of columns in schedule file: Got {len(sheet.columns)}.")
        return False

    else:
        ret = True

        for i, v in enumerate(columns):
            if sheet.columns[i].title != v:
                print(f"   Mismatch schedule column name for col {i+1}. Got {sheet.columns[i].title}. Expect {v}.")
                ret = False

        return ret

def check_parent_row(*, client, sheet, rowIdx, doFixes, title):
    formulas = {7: '=NETWORKDAYS([Baseline Start]@row, [Baseline Finish]@row)',
                8:  '=SUM(CHILDREN())',
                12: '=SUM(CHILDREN())',
                13: '=SUM(CHILDREN())' }

    # These Columns SHould Have No Value
    noValue = set([9, 15, 16])

    # Preserve Values, but apply formatting
    noChange = set([14])

    if rowIdx >= len(sheet.rows):
        return

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    # Check first Cell of first row
    if title is not None:
        form = formatMajor

        if row.cells[0].value != title or row.cells[0].format != form[0]:
            print(f"   Incorrect format in row {rowIdx+1} cell 1 in schedule file. Got '{row.cells[0].format}' Exp '{form[0]}'")
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[0].id
            new_cell.value = title
            new_cell.format = form[0]
            new_cell.strict = False
            new_row.cells.append(new_cell)

    else:
        form = formatMinor

    new_row.format = form[0]

    for i in range(1,len(row.cells)):
        if i in formulas:

            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i] or (form[i] is not None and row.cells[i].format != form[i]):

                if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i]:
                    print(f"   Incorrect formula in row {rowIdx+1} cell {i+1} in schedule file. Got '{row.cells[i].formula}' Exp '{formulas[i]}'")

                if (form[i] is not None and row.cells[i].format != form[i]):
                    print(f"   Incorrect format in row {rowIdx+1} cell {i+1} in schedule file. Got '{row.cells[i].format}' Exp '{form[i]}'")

                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = formulas[i]

                if form[i] is not None:
                    new_cell.format = form[i]

                new_cell.strict = False
                new_row.cells.append(new_cell)

        if i in noValue:
            if (form[i] is not None and row.cells[i].format != form[i]) or (row.cells[i].value is not None and row.cells[i].value != ''):

                if (row.cells[i].value is not None and row.cells[i].value != ''):
                    print(f"   Incorrect value in row {rowIdx+1} cell {i+1} in schedule file. Got '{row.cells[i].value}' Exp ''.")

                if (form[i] is not None and row.cells[i].format != form[i]):
                    print(f"   Incorrect format in row {rowIdx+1} cell {i+1} in schedule file. Got '{row.cells[i].format}' Exp '{form[i]}'")

                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.value = ''

                if form[i] is not None:
                    new_cell.format = form[i]

                new_cell.strict = False
                new_row.cells.append(new_cell)

        if i in noChange:
            if form[i] is not None and row.cells[i].format != form[i]:
                print(f"   Incorrect format in row {rowIdx+1} cell {i+1} in schedule file. Got '{row.cells[i].format}' Exp '{form[i]}'")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id

                if row.cells[i].value is None:
                    new_cell.value = ''
                else:
                    new_cell.value = row.cells[i].value
                new_cell.format = form[i]
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
             14 : '0',
             15 : '0',
             16 : '0',
             18 : 'No',
           }

    # Preserve Values, but apply formatting
    noChange = set([9, 17, 19])

    if rowIdx >= len(sheet.rows):
        return

    row = sheet.rows[rowIdx]
    form = formatTask

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for i in range(len(row.cells)):
        if i in formulas:

            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i] or (form[i] is not None and row.cells[i].format != form[i]):

                if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i]:
                    print(f"   Incorrect formula in row {rowIdx+1} cell {i+1} in schedule file. Got '{row.cells[i].formula}' Exp '{formulas[i]}'.")

                if (form[i] is not None and row.cells[i].format != form[i]):
                    print(f"   Incorrect format in row {rowIdx+1} cell {i+1} in schedule file. Got '{row.cells[i].format}' Exp '{form[i]}'")

                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = formulas[i]
                new_cell.strict = False
                new_cell.format = form[i]
                new_row.cells.append(new_cell)

        if i in init:
            if row.cells[i].value is None or row.cells[i].value == '' or (form[i] is not None and row.cells[i].format != form[i]):

                if row.cells[i].value is None or row.cells[i].value == '':
                    print(f"   Incorrect default value in row {rowIdx+1} cell {i+1} in schedule file.")

                if (form[i] is not None and row.cells[i].format != form[i]):
                    print(f"   Incorrect format in row {rowIdx+1} cell {i+1} in schedule file. Got '{row.cells[i].format}' Exp '{form[i]}'")

                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.format = form[i]
                new_cell.strict = False

                if row.cells[i].value is not None and row.cells[i].value != '':
                    new_cell.value = row.cells[i].value
                else:
                    new_cell.value = init[i]

                new_row.cells.append(new_cell)

        if i in noChange:
            if form[i] is not None and row.cells[i].format != form[i]:
                print(f"   Incorrect format in row {rowIdx+1} cell {i+1} in schedule file. Got '{row.cells[i].format}' Exp '{form[i]}'")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id

                if row.cells[i].value is not None and row.cells[i].value != '':
                    new_cell.value = row.cells[i].value
                else:
                    new_cell.value = ''

                new_cell.format = form[i]
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to schedule row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_parent_links(*, client, sheet, rowIdx, laborRows, laborSheet, doFixes):
    links = { 0: 0 }

    if rowIdx >= len(sheet.rows):
        return

    row = sheet.rows[rowIdx]
    form = formatMinor

    # Need re-link
    relink = set()

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    # First check format
    for i in range(len(row.cells)):
        if i in links:

            if form[i] is not None and row.cells[i].format != form[i]:
                print(f"   Incorrect format in row {row.row_number} cell {i+1} in schedule file. Got '{row.cells[i].format}' Exp '{form[i]}'")
                relink.add(i)
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = row.cells[i].column_id
                new_cell.value = smartsheet.models.ExplicitNull()
                new_cell.format = form[i]
                new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to schedule row {row.row_number}.")
        client.Sheets.update_rows(sheet.id, [new_row])

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for i in range(len(row.cells)):
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

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to schedule row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_links(*, client, sheet, rowIdx, laborRows, laborSheet, doFixes):
    links = { 8: 2, 20: 20 }

    if rowIdx >= len(sheet.rows):
        return

    row = sheet.rows[rowIdx]
    form = formatTask

    # Need re-link
    relink = set()

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    # First check format
    for i in range(1,len(row.cells)):
        if i in links:

            if form[i] is not None and row.cells[i].format != form[i]:
                print(f"   Incorrect format in row {row.row_number} cell {i+1} in schedule file. Got '{row.cells[i].format}' Exp '{form[i]}'")
                relink.add(i)
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = row.cells[i].column_id
                new_cell.value = smartsheet.models.ExplicitNull()
                new_cell.format = form[i]
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
            print(f"   Incorrect schedule link for row {rowIdx+1} column {i+1}.")
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = row.cells[i].column_id
            new_cell.value = ''

            if form[i] is not None:
                new_cell.format = form[i]

            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to schedule row {rowIdx+1}.")
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
        print("   Deleting stale rows.")
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
            check_parent_links(client=client, sheet=sheet, rowIdx=i+1, laborRows=laborRows, laborSheet=laborSheet, doFixes=doFixes)
        else:
            check_task_row(client=client, sheet=sheet, rowIdx=i+1, doFixes=doFixes)
            check_task_links(client=client, sheet=sheet, rowIdx=i+1, laborRows=laborRows, laborSheet=laborSheet, doFixes=doFixes)

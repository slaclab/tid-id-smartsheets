#-----------------------------------------------------------------------------
# Title      : Manipulate Budget Sheet
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

# Set formats for columns looking at major, minor and task rows
#
# https://smartsheet-platform.github.io/api-docs/#formatting
#
# Colors 31 = Dark Blue
#        26 = Dark Gray
#        23 = Blue
#        18 - Gray
#        22 = Green
#           = White

formatMajor = [ ",,1,,,,,,2,31,,,,,,1,",     ",,1,,,,,,2,31,,,,,,1,",     ",,1,,,,,,2,31,,,,,,1,",     ",,1,,,,,,2,31,,,,,,1,",     ",,1,,,,,,2,31,,,,,,1,",
                ",,1,,,,,,2,31,,,,,,1,",     ",,1,,,,,,2,31,,13,2,1,2,,", ",,1,,,,,,2,31,,13,2,1,2,,", ",,1,,,,,,2,31,,,,,,1,",     ",,1,,,,,,2,31,,,,,,1,",
                ",,1,,,,,,2,31,,,,,,1,",     ",,1,,,,,,2,31,,,,,,1,",     ",,1,,,,,,2,31,,13,2,1,2,,", ",,1,,,,,,2,31,,13,2,1,2,,", ",,1,,,,,,2,31,,,,,,1,",
                ",,1,,,,,,2,31,,13,2,1,2,,", ",,1,,,,,,2,31,,13,2,1,2,,", ",,1,,,,,,2,31,,13,2,1,2,,", ",,1,,,,,,2,31,,,0,,,1,",    ",,1,,,,,,2,31,,13,2,1,2,,",
                ",,1,,,,,,2,31,,13,2,1,2,,", ",,1,,,,,,2,31,,13,2,1,2,,", ",,1,,,,,,2,31,,13,2,1,2,,", ",,1,,,,,,2,31,,,,,,1," ]

formatMinor = [ ",,,,,,,,,23,,,,,,,",      ",,,,,,,,,23,,,,,,,",      ",,,,,,,,,23,,,,,,,",      ",,,,,,,,,23,,,,,,,",      ",,,,,,,,,23,,,,,,,",
                ",,,,,,,,,23,,,,,,,",      ",,,,,,,,,23,,13,2,1,2,,", ",,,,,,,,,23,,13,2,1,2,,", ",,,,,,,,,23,,,,,,,",      ",,,,,,,,,23,,,,,,,",
                ",,,,,,,,,23,,,,,,,",      ",,,,,,,,,23,,,,,,1,",     ",,,,,,,,,23,,13,2,1,2,,", ",,,,,,,,,23,,13,2,1,2,,", ",,,,,,,,,23,,,,,,,",
                ",,,,,,,,,23,,13,2,1,2,,", ",,,,,,,,,23,,13,2,1,2,,", ",,,,,,,,,23,,13,2,1,2,,", ",,,,,,,,,23,,,0,,,,",     ",,,,,,,,,23,,13,2,1,2,,",
                ",,,,,,,,,23,,13,2,1,2,,", ",,,,,,,,,23,,13,2,1,2,,", ",,,,,,,,,23,,13,2,1,2,,", ",,,,,,,,,23,,,,,,,"]

formatMAndS = [ ",,,,,,,,,2,,,,,,1,",      ",,,,,,,,,2,,,,,,1,",      ",,,,,,,,,2,,,,,,1,",      ",,,,,,,,,2,,13,2,1,2,,",  ",,,,,,,,,2,,,,,,,",
                ",,,,,,,,,18,,,,,3,,",     ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,26,,,,,,,",      ",,,,,,,,,26,,,,,,,",
                ",,,,,,,,,26,,,,,,,",      ",,,,,,,,,22,,,,,3,,",     ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,26,,,,,,,",
                ",,,,,,,,,22,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,,0,,3,,",    ",,,,,,,,,18,,13,2,1,2,,",
                ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,2,,,,,,," ]

formatTask  = [ ",,,,,,,,,2,,,,,,,",       ",,,,,,,,,2,,,,,,,",       ",,,,,,,,,2,,,,,,,",       ",,,,,,,,,2,,13,2,1,2,,",  ",,,,,,,,,2,,,,,,,",
                ",,,,,,,,,18,,,,,3,,",     ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,,,,,,",      ",,,,,,,,,18,,,0,,,,",
                ",,,,,,,,,18,,,,,,,",      ",,,,,,,,,18,,,,,3,,",     ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,,,,,,",
                ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,,0,,3,,",    ",,,,,,,,,18,,13,2,1,2,,",
                ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,18,,13,2,1,2,,", ",,,,,,,,,2,,,,,,," ]

Columns = ['Category and Item',  # 0
           'Item Notes',  # 1
           'Budgeted Labor Hours Or Qty',  # 2
           'Cost per Item Fully Burdened', # 3
           'Risk Factor',  # 4
           'Contingency Factor',  # 5
           'Total Budget',  # 6
           'Budget With Contingency',  # 7
           'Planned Duration From Schedule (days)',      # 8
           'Planned Duration With Contingency  (days)',  # 9
           'Assigned To From Schedule',  # 10
           'Percent Complete From Schedule Or Local M&S', # 11
           'Earned Value',  # 12
           'Earned Value With Contingency',    # 13
           'Actual Hours From Schedule',       # 14
           'Actual Cost Computed Or Real M&S', # 15
           'Budget Variance',  # 16
           'Budget Variance With Contingency',  # 17
           'Planned % Complete From Schedule',  # 18  # Rename
           'Planned Earned Value',  # 19  # Rename
           'Planned Earned Value With Contingency', # 20 # Add
           'Schedule Variance',  # 21 # Add
           'Schedule Variance With Contingency', # 22 # Add
           'PA Number']  # 23

def fix_structure(*, client, sheet):

    return False

    if len(sheet.columns) != 21:
        print(f"   Wrong number of columns in budget file, could not fix: Got {len(sheet.columns)}.")
        return False

    # Modify Column Names
    col18 = smartsheet.models.Column({'title': Columns[21],
                                      'type': 'TEXT_NUMBER',
                                      'index': 18})

    col19 = smartsheet.models.Column({'title': Columns[22],
                                      'type': 'TEXT_NUMBER',
                                      'index': 19})

    client.Sheets.update_column(sheet.id, sheet.columns[18].id, col18)
    client.Sheets.update_column(sheet.id, sheet.columns[19].id, col19)

    # Add new columns
    col20 = smartsheet.models.Column({'title': Columns[18],
                                      'type': 'TEXT_NUMBER',
                                      'index': 18})

    col21 = smartsheet.models.Column({'title': Columns[19],
                                      'type': 'TEXT_NUMBER',
                                      'index': 18})

    col22 = smartsheet.models.Column({'title': Columns[20],
                                      'type': 'TEXT_NUMBER',
                                      'index': 18})

    client.Sheets.add_columns(sheet.id, [col20, col21, col22])

    return True


def check_structure(*, sheet):

    # Check column count
    if len(sheet.columns) != len(Columns):
        print(f"   Wrong number of columns in budget file: Got {len(sheet.columns)} Expect {len(Columns)}.")
        return False

    else:
        ret = True

        for i,v in enumerate(Columns):
            if sheet.columns[i].title != v:
                print(f"   Mismatch budget column name for col {i+1}. Got {sheet.columns[i].title}. Expect {v}.")
                ret = False

        return ret


def check_parent_row(*, client, sheet, rowIdx, sumCols, titles, doFixes):
    sumChildrenValue = '=SUM(CHILDREN())'
    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    if len(titles) == 0:
        startIdx = 1
        form = formatMinor
    else:
        startIdx = len(titles)
        form = formatMajor

    for i in range(len(titles)):
        if row.cells[i].value != titles[i] or row.cells[i].format != form[i]:

            if row.cells[i].value != titles[i]:
                print(f"   Incorrect title in row {rowIdx+1} cell {i+1} in budget file. Got '{row.cells[i].value}' Expect '{titles[i]}'.")

            if row.cells[i].format != form[i]:
                print(f"   Incorrect format in row {rowIdx+1} cell {i+1} in budget file. Got '{row.cells[i].format}' Expect '{form[i]}'.")

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[i].id
            new_cell.value = titles[i]
            new_cell.format = form[i]
            new_cell.strict = False
            new_row.cells.append(new_cell)

    # Check format in earlier columns
    for i in range(len(titles), startIdx):
        if row.cells[i].format != form[i]:
            print(f"   Incorrect format in row {rowIdx+1} cell {i+1} in budget file.. Got '{row.cells[i].format}' Expect '{form[i]}'")
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[i].id
            new_cell.value = row.cells[i].value
            new_cell.format = form[i]
            new_cell.strict = False
            new_row.cells.append(new_cell)

    for i in range(startIdx, len(row.cells)):

        if i in sumCols:
            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != sumChildrenValue or row.cells[i].format != form[i]:

                if not hasattr(row.cells[i],'formula') or row.cells[i].formula != sumChildrenValue:
                    print(f"   Incorrect formula in row {rowIdx+1} cell {i+1} in budget file. Expect '{sumChildrenValue}'")

                if row.cells[i].format != form[i]:
                    print(f"   Incorrect format in row {rowIdx+1} cell {i+1} in budget file. Got '{row.cells[i].format}' Expect '{form[i]}'")

                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = sumChildrenValue
                new_cell.strict = False
                new_cell.format = form[i]
                new_row.cells.append(new_cell)

        elif row.cells[i].format != form[i] or (row.cells[i].value is not None and row.cells[i].value != ''):

            if row.cells[i].format != form[i]:
                print(f"   Incorrect format in row {rowIdx+1} cell {i+1} in budget file. Got '{row.cells[i].format}' Expect '{form[i]}'")

            if (row.cells[i].value is not None and row.cells[i].value != ''):
                print(f"   Incorrect value in row {rowIdx+1} cell {i+1} in budget file. Got '{row.cells[i].value}' Exp ''")

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[i].id
            new_cell.value = ''
            new_cell.strict = False
            new_cell.format = form[i]
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to budget row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_row(*, client, sheet, rowIdx, inMS, doFixes):
    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    # Row formulas
    rowVals = {5:  '=IF([Risk Factor]@row = "Low (5% Contingency)", 1.05, '
                   'IF([Risk Factor]@row = "Medium (10% Contingency)", 1.1, '
                   'IF([Risk Factor]@row = "Med-High (25% Contingency)", 1.25, '
                   'IF([Risk Factor]@row = "High (50% Contingency)", 1.5))))',
               6:  '=[Budgeted Labor Hours Or Qty]@row * [Cost per Item Fully Burdened]@row',
               7:  '=[Contingency Factor]@row * [Total Budget]@row',
               12: '=[Total Budget]@row * [Percent Complete From Schedule Or Local M&S]@row',
               13: '=[Earned Value]@row * [Contingency Factor]@row',
               16: '=[Actual Cost Computed Or Real M&S]@row - [Earned Value]@row',
               17: '=[Actual Cost Computed Or Real M&S]@row - [Earned Value With Contingency]@row',
               19: '=[Total Budget]@row * [Planned % Complete From Schedule]@row',
               20: '=[Planned Earned Value]@row * [Contingency Factor]@row',
               21: '=[Planned Earned Value]@row - [Earned Value]@row',
               22: '=[Planned Earned Value With Contingency]@row - [Earned Value With Contingency]@row',
              }

    # Empty cells
    empty = []

    # Cells that should have defaults
    defV  = []

    # Preserve Values, but apply formatting
    noChange = [0, 1, 2, 3, 4, 23]

    if inMS:
        empty.append(8)
        empty.append(9)
        empty.append(10)
        empty.append(14)
        defV.append(11)
        defV.append(15)
        rowVals[18] = '=[Percent Complete From Schedule Or Local M&S]@row'
        form = formatMAndS
    else:
        rowVals[9]  = '=[Contingency Factor]@row * [Planned Duration From Schedule (days)]@row'
        rowVals[15] = '=[Cost per Item Fully Burdened]@row * [Actual Hours From Schedule]@row'
        form = formatTask

    for k,v in rowVals.items():
        if ((not hasattr(row.cells[k],'formula')) or row.cells[k].formula != v) or row.cells[k].format != form[k]:

            if ((not hasattr(row.cells[k],'formula')) or row.cells[k].formula != v):
                print(f"   Incorrect formula in row {rowIdx+1} cell {k+1} in budget file.")

            if row.cells[k].format != form[k]:
                print(f"   Incorrect format in row {rowIdx+1} cell {k+1} in budget file. Got '{row.cells[k].format}' Expect '{form[k]}'")

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[k].id
            new_cell.formula = v
            new_cell.format = form[k]
            new_cell.strict = False
            new_row.cells.append(new_cell)

    for k in empty:
        if (row.cells[k].value is not None and row.cells[k].value != '') or row.cells[k].format != form[k]:

            if (row.cells[k].value is not None and row.cells[k].value != ''):
                print(f"   Bad value in row {rowIdx+1} col {k+1} in budget file. Got '{row.cells[k].value}' Exp ''")

            if row.cells[k].format != form[k]:
                print(f"   Incorrect format in row {rowIdx+1} cell {k+1} in budget file. Got '{row.cells[k].format}' Expect '{form[k]}'")

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[k].id
            new_cell.value = ''
            new_cell.format = form[k]
            new_cell.strict = False
            new_row.cells.append(new_cell)

    for k in defV:
        if row.cells[k].value is None or row.cells[k].value == '' or row.cells[k].format != form[k]:

            if row.cells[k].value is None or row.cells[k].value == '':
                print(f"   Missing default value in row {rowIdx+1} col {k+1} in budget file.")

            if row.cells[k].format != form[k]:
                print(f"   Incorrect format in row {rowIdx+1} cell {k+1} in budget file. Got '{row.cells[k].format}' Expect '{form[k]}'")

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[k].id
            new_cell.value = '0'
            new_cell.format = form[k]
            new_cell.strict = False

            if row.cells[k].value is not None and row.cells[k].value != '':
                new_cell.value = row.cells[k].value
            else:
                new_cell.value = '0'

            new_row.cells.append(new_cell)

    for k in noChange:
        if form[k] is not None and row.cells[k].format != form[k]:
            print(f"   Incorrect format in row {rowIdx+1} cell {k+1} in budget file. Got '{row.cells[k].format}' Expect '{form[k]}'")
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[k].id

            if row.cells[k].value is None:
                new_cell.value = ''
            else:
                new_cell.value = row.cells[k].value
            new_cell.format = form[k]
            new_cell.strict = False
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to budget row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_links(*, client, sheet, laborRows, scheduleSheet, doFixes):
    links = { 8: 7,
             10: 9,
             11: 14,
             14: 13,
             18: 12}

    # Need re-link
    relink = set()

    form = formatTask

    # First check format
    for rowData in laborRows:
        if not rowData['parent']:
            row = rowData['data']

            # Setup row update structure just in case
            new_row = smartsheet.models.Row()
            new_row.id = row.id

            for i in range(1,len(row.cells)):
                if i in links:

                    if row.cells[i].format != form[i]:
                        print(f"   Incorrect format in row {row.row_number} cell {i+1} in budget file. Got '{row.cells[i].format}' Expect '{form[i]}'")
                        relink.add(i)
                        new_cell = smartsheet.models.Cell()
                        new_cell.column_id = row.cells[i].column_id
                        new_cell.value = smartsheet.models.ExplicitNull()
                        new_cell.format = form[i]
                        new_row.cells.append(new_cell)

            if doFixes and len(new_row.cells) != 0:
                print(f"   Applying fixes to budget row {row.row_number}.")
                client.Sheets.update_rows(sheet.id, [new_row])

    for rowData in laborRows:
        if not rowData['parent']:
            row = rowData['data']

            if rowData['link'] is None:
                print(f"   Missing budget link for row {row.row_number}.")
            else:

                # Setup row update structure just in case
                new_row = smartsheet.models.Row()
                new_row.id = row.id

                for i in range(1,len(row.cells)):
                    if i in links:
                        rowIdTar = rowData['link'].id
                        colIdTar = rowData['link'].cells[links[i]].column_id
                        shtIdTar = scheduleSheet.id

                        if i in relink or row.cells[i].link_in_from_cell is None or \
                            row.cells[i].link_in_from_cell.row_id != rowIdTar or \
                            row.cells[i].link_in_from_cell.column_id != colIdTar or \
                            row.cells[i].link_in_from_cell.sheet_id != shtIdTar:

                            print(f"   Incorrect budget link for row {row.row_number} column {i+1}.")

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
                    print(f"   Applying fixes to budget row {row.row_number}.")
                    client.Sheets.update_rows(sheet.id, [new_row])


def check(*, client, sheet, doFixes):
    labor = []

    check_parent_row(client=client,
                  sheet=sheet,
                  rowIdx=0,
                  sumCols=set([6, 7, 12, 13, 15, 16, 17, 19, 20, 21, 22]),
                  titles=['Total Project Budget'],
                  doFixes=doFixes)

    check_parent_row(client=client,
                  sheet=sheet,
                  rowIdx=1,
                  sumCols=set([6, 7, 12, 13, 15, 16, 17, 19, 20, 21, 22]),
                  titles=['M&S Total',navigate.OVERHEAD_NOTE],
                  doFixes=doFixes)


    # First walk through the rows and create a parent list
    parents = set()
    for rowIdx in range(2,len(sheet.rows)):
        parents.add(sheet.rows[rowIdx].parent_id)

    inMS = True

    for rowIdx in range(2,len(sheet.rows)):

        # Start of labor section
        if sheet.rows[rowIdx].cells[0].value == 'Labor':
            check_parent_row(client=client,
                          sheet=sheet,
                          rowIdx=rowIdx,
                          sumCols=set([2, 6, 7, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22]),
                          titles=['Labor', navigate.LABOR_RATE_NOTE],
                          doFixes=doFixes)

            inMS = False

        # This is a parent row
        elif sheet.rows[rowIdx].id in parents:

            # M&S Parent Rows
            if inMS:
                check_parent_row(client=client,
                                 sheet=sheet,
                                 rowIdx=rowIdx,
                                 sumCols=set([6, 7, 12, 13, 15, 16, 17, 19, 20, 21, 22]),
                                 titles=[],
                                 doFixes=doFixes)

            # Labor Parent Rows
            else:

                labor.append({'data': sheet.rows[rowIdx], 'parent' : True, 'link' : None})

                check_parent_row(client=client,
                                 sheet=sheet,
                                 rowIdx=rowIdx,
                                 sumCols=set([2, 6, 7, 12, 13, 14, 15, 16, 17, 19, 20, 21, 22]),
                                 titles=[],
                                 doFixes=doFixes)

        # Skip rows with missing text
        elif hasattr(sheet.rows[rowIdx].cells[0],'value') and sheet.rows[rowIdx].cells[0].value is not None:

            if not inMS:
                labor.append({'data': sheet.rows[rowIdx], 'parent' : False, 'link' : None})

            check_task_row(client=client,
                           sheet=sheet,
                           rowIdx=rowIdx,
                           inMS=inMS,
                           doFixes=doFixes)

    return labor

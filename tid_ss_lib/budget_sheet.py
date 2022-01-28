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

# TODO:
#    Remove uneccessary links, particulary in parent rows
#    Color format parent rows


def check_structure(*, sheet):

    columns = ['Category and Item',
               'Item Notes',
               'Budgeted Labor Hours Or Qty',
               'Cost per Item Fully Burdened',
               'Risk Factor',
               'Contingency Factor',
               'Total Budget',
               'Budget With Contingency',
               'Planned Duration From Schedule (days)',
               'Planned Duration With Contingency  (days)',
               'Assigned To From Schedule',
               'Percent Complete From Schedule Or Local M&S',
               'Earned Value',
               'Earned Value With Contingency',
               'Actual Hours From Schedule',
               'Actual Cost Computed Or Real M&S',
               'Budget Variance',
               'Budget Variance With Contingency',
               'Duration Variance From Schedule',
               'Duration Variance With Contingency',
               'PA Number']

    # Check column count
    if len(sheet.columns) != len(columns):
        print(f"   Wrong number of columns in budget file: Got {len(sheet.columns)}.")
        return False

    else:
        ret = True

        for i,v in enumerate(columns):
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

    for num in range(len(titles)):
        if row.cells[num].value != titles[num]:
            print(f"   Incorrect row {rowIdx+1}, col {num+1} value in budget file. Got '{row.cells[num].value}'. Expected '{titles[num]}'.")
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[num].id
            new_cell.value = titles[num]
            new_cell.strict = False
            new_row.cells.append(new_cell)

        # Check shading
        #print(row.cells[num])

    if len(titles) == 0:
        startIdx = 1
    else:
        startIdx = len(titles)

    for i in range(startIdx,len(row.cells)):

        if i in sumCols:
            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != sumChildrenValue:
                print(f"   Invalid value in row {rowIdx+1} cell {i+1} in budget file. Expected '{sumChildrenValue}'. Got '{row.cells[i].formula}'.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = sumChildrenValue
                new_cell.strict = False
                new_row.cells.append(new_cell)

        else:
            if row.cells[i].value is not None:
                print(f"   Invalid value in row {rowIdx+1} cell {i+1} in budget file. Expected ''. Got '{row.cells[i].value}'.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.value = ''
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_row(*, client, sheet, rowIdx, inMS, doFixes):
    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    # Row formulas
    rowVals = {5: '=IF([Risk Factor]@row = "Low (5% Contingency)", 1.05, '
                  'IF([Risk Factor]@row = "Medium (10% Contingency)", 1.1, '
                  'IF([Risk Factor]@row = "Med-High (25% Contingency)", 1.25, '
                  'IF([Risk Factor]@row = "High (50% Contingency)", 1.5))))',
               6: '=[Budgeted Labor Hours Or Qty]@row * [Cost per Item Fully Burdened]@row',
               7: '=[Contingency Factor]@row * [Total Budget]@row',

               12: '=[Total Budget]@row * [Percent Complete From Schedule Or Local M&S]@row',
               13: '=[Earned Value]@row * [Contingency Factor]@row',
               15: '=[Cost per Item Fully Burdened]@row * [Actual Hours From Schedule]@row',  # Not M&S
               16: '=[Actual Cost Computed Or Real M&S]@row - [Earned Value]@row',
               17: '=[Actual Cost Computed Or Real M&S]@row - [Earned Value With Contingency]@row',
               19: '=[Duration Variance From Schedule]@row - (([Contingency Factor]@row - 1) * [Planned Duration From Schedule (days)]@row)'} # Not M&S

    for k,v in rowVals.items():
        if not (inMS and (k == 15 or k == 19)):
            if (not hasattr(row.cells[k],'formula')) or row.cells[k].formula != v:
                print(f"   Invalid value in row {rowIdx+1} cell {k+1} in budget file. Expected '{v}'. Got '{row.cells[k].formula}'.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[k].id
                new_cell.formula = v
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_links(*, client, sheet, laborRows, scheduleSheet, doFixes):
    links = { 8: 7,
             10: 9,
             11: 14,
             14: 13,
             18: 12}

    for rowData in laborRows:
        if not rowData['parent']:
            row = rowData['data']

            # Setup row update structure just in case
            new_row = smartsheet.models.Row()
            new_row.id = row.id

            for k,v in links.items():
                if rowData['link'] is None:
                    print(f"   Missing budget link for row {row.row_number} column {k+1}.")
                else:

                    rowIdTar = rowData['link'].id
                    colIdTar = rowData['link'].cells[v].column_id
                    shtIdTar = scheduleSheet.id

                    if row.cells[k].link_in_from_cell is None or \
                        row.cells[k].link_in_from_cell.row_id != rowIdTar or \
                        row.cells[k].link_in_from_cell.column_id != colIdTar or \
                        row.cells[k].link_in_from_cell.sheet_id != shtIdTar:

                        print(f"   Incorrect budget link for row {row.row_number} column {k+1}.")

                        cell_link = smartsheet.models.CellLink()
                        cell_link.sheet_id = shtIdTar
                        cell_link.row_id = rowIdTar
                        cell_link.column_id = colIdTar

                        new_cell = smartsheet.models.Cell()
                        new_cell.column_id = row.cells[k].column_id
                        new_cell.value = smartsheet.models.ExplicitNull()
                        new_cell.link_in_from_cell = cell_link
                        new_row.cells.append(new_cell)

            if doFixes and len(new_row.cells) != 0:
                print(f"   Applying fixes to row {row.row_number}.")
                client.Sheets.update_rows(sheet.id, [new_row])


def check(*, client, sheet, doFixes):
    labor = []

    check_parent_row(client=client,
                  sheet=sheet,
                  rowIdx=0,
                  sumCols=set([6, 7, 12, 13, 15, 16, 17, 18, 19]),
                  titles=['Total Project Budget'],
                  doFixes=doFixes)

    check_parent_row(client=client,
                  sheet=sheet,
                  rowIdx=1,
                  sumCols=set([6, 7, 12, 13, 15, 16, 17]),
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
                          sumCols=set([2, 6, 7, 12, 13, 14, 15, 16, 17, 18, 19]),
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
                                 sumCols=set([6, 7, 12, 13, 15, 16, 17]),
                                 titles=[],
                                 doFixes=doFixes)

            # Labor Parent Rows
            else:

                labor.append({'data': sheet.rows[rowIdx], 'parent' : True, 'link' : None})

                check_parent_row(client=client,
                                 sheet=sheet,
                                 rowIdx=rowIdx,
                                 sumCols=set([2, 6, 7, 12, 13, 14, 15, 16, 17, 18, 19]),
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


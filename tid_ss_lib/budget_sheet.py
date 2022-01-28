
import smartsheet


def check_parent_row(*, client, sheet, rowIdx, sumCols, titles):
    sumChildrenValue = '=SUM(CHILDREN())'
    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for num in range(len(titles)):
        if row.cells[num].value != titles[num]:
            print(f"   Incorrect row {rowIdx}, col {num} value in budget file. Got '{row.cells[num].value}'. Expected '{titles[num]}'. Fixing.")
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[num].id
            new_cell.value = titles[num]
            new_cell.strict = False
            new_row.cells.append(new_cell)

    if len(titles) == 0:
        startIdx = 1
    else:
        startIdx = len(titles)

    for i in range(startIdx,21):

        if i in sumCols:
            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != sumChildrenValue:
                print(f"   Invalid value in row {rowIdx} cell {i} in budget file. Expected '{sumChildrenValue}'. Got '{row.cells[i].formula}'. Fixing")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = sumChildrenValue
                new_cell.strict = False
                new_row.cells.append(new_cell)

        else:
            if row.cells[i].value is not None:
                print(f"   Invalid value in row {rowIdx} cell {i} in budget file. Expected ''. Got '{row.cells[i].value}'. Fixing.")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.value = ''
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_row(*, client, sheet, rowIdx, inMS):
    row = sheet.rows[rowIdx]

    # Skip rows with missing text
    if (not hasattr(row.cells[0],'value')) or row.cells[0].value is None:
        return

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
                print(f"   Invalid value in row {rowIdx} cell {k} in budget file. Expected '{v}'. Got '{row.cells[k].formula}'. Fixing")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[k].id
                new_cell.formula = v
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check(*, client, sheet):

    check_parent_row(client=client,
                  sheet=sheet,
                  rowIdx=0,
                  sumCols=set([6, 7, 12, 13, 15, 16, 17, 18, 19]),
                  titles=['Total Project Budget'])

    check_parent_row(client=client,
                  sheet=sheet,
                  rowIdx=1,
                  sumCols=set([6, 7, 12, 13, 15, 16, 17]),
                  titles=['M&S Total','12.25% Overhead'])


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
                          titles=['Labor', 'SLAC Labor Rate FY21 (loaded): $274.63; Slac Tech Rate FY21 (loaded): $163.47'])

            inMS = False

        # This is a parent row
        elif sheet.rows[rowIdx].id in parents:

            # M&S Parent Rows
            if inMS:
                check_parent_row(client=client,
                                 sheet=sheet,
                                 rowIdx=rowIdx,
                                 sumCols=set([6, 7, 12, 13, 15, 16, 17]),
                                 titles=[])

            # Labor Parent Rows
            else:
                check_parent_row(client=client,
                                 sheet=sheet,
                                 rowIdx=rowIdx,
                                 sumCols=set([6, 7, 12, 13, 14, 15, 16, 17, 18, 19]),
                                 titles=[])

        else:
            check_task_row(client=client,
                           sheet=sheet,
                           rowIdx=rowIdx,
                           inMS=inMS)



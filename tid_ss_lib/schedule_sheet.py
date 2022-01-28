
import smartsheet

def check_parent_row(*, client, sheet, rowIdx):
    formulas = [  None, None, None, None, None, None, None,
                '=NETWORKDAYS([Baseline Start]@row, [Baseline Finish]@row)',
                None, None, None, None,
                '=SUM(CHILDREN())',
                '=SUM(CHILDREN())',
                None, None, None]

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for i in range(len(formulas)):

        if formulas[i] is not None:

            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i]:
                print(f"   Invalid value in row {rowIdx} cell {i} in schedule file. Expected '{formulas[i]}'. Got '{row.cells[i].formula}'. Fixing")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = formulas[i]
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_row(*, client, sheet, rowIdx):
    formulas = [  None, None, None, None, None, None, None,
                '=NETWORKDAYS([Baseline Start]@row, [Baseline Finish]@row)',
                None, None,
                '=([Planned Labor Hours From Budget]@row / 8) / [Baseline Duration (days)]@row',
                None,
                '=Duration@row - [Baseline Duration (days)]@row',
                None, None, None, None]

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for i in range(len(formulas)):

        if formulas[i] is not None:

            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i]:
                print(f"   Invalid value in row {rowIdx} cell {i} in schedule file. Expected '{formulas[i]}'. Got '{row.cells[i].formula}'. Fixing")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = formulas[i]
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_links(*, client, sheet, rowIdx):
    links = [ None, None, None, None, None, None, None,
              '=NETWORKDAYS([Baseline Start]@row, [Baseline Finish]@row)',
              None, None, None, None,
              '=SUM(CHILDREN())',
              '=SUM(CHILDREN())',
              None, None, None]


# Find stale or broken links to budget file
def check_broken(*, client, sheet, laborSheet):

    for rowIdx in range(1,len(sheet.rows)):
        row = sheet.rows[rowIdx]

        #if (not hasattr(row.cells[0],'linkInFromCell')) or row.cells[0].linkInFromCell.status == 'Broken':
        if row.cells[0].link_in_from_cell is None or row.cells[0].link_in_from_cell.status != 'OK':
            print(f"   Found stale row {rowIdx} in schedule file. Fixing")
            delList.append(row.id)

    if len(delList) != 0:
        pass
        client.Sheets.delete_rows(sheet.id, delList)


# Check matching rows to budget file
def check_rows(*, client, sheet, laborRows, laborSheet):
    lastId = sheet.rows[0].id
    rowIdx = 1

    for i in range(len(laborRows)):
        if rowIdx < len(sheet.rows) and sheet.rows[rowIdx].cells[0].link_in_from_cell.row_id == laborRows[i]['data'].id:
            lastId = sheet.rows[rowIdx].id
            rowIdx += 1
        else:

            print(f"   Adding new schedule row at position {rowIdx}, budget row = {i}, parent = {laborRows[i]['parent']}")

            newRow = smartsheet.models.Row()
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



def check(*, client, sheet, laborRows, laborSheet):

    check_broken(client=client, sheet=sheet)
    sheet = client.Sheets.get_sheet(sheet.id)

    while check_rows(client=client, sheet=sheet, laborRows=laborRows, laborSheet=laborSheet):
        sheet = client.Sheets.get_sheet(sheet.id)

    # Process top row
    check_parent_row(client=client, sheet=sheet, rowIdx=0)

    # Process rows
    for i in range(len(laborRows)):

        if laborRows[i]['parent']:
            check_parent_row(client=client, sheet=sheet, rowIdx=i+1)
        else:
            check_task_row(client=client, sheet=sheet, rowIdx=i+1)
            # Check links




        # Start of labor section
        #if sheet.rows[rowIdx].cells[0].value == 'Labor':
        #    check_parent_row(client=client,
        #                  sheet=sheet,
        #                  rowIdx=rowIdx,
        #                  sumCols=set([2, 6, 7, 12, 13, 14, 15, 16, 17, 18, 19]),
        #                  titles=['Labor', 'SLAC Labor Rate FY21 (loaded): $274.63; Slac Tech Rate FY21 (loaded): $163.47'])

        #    inMS = False






#    check_parent_row(client=client,
#                  sheet=sheet,
#                  rowIdx=0,
#                  sumCols=set([6, 7, 12, 13, 15, 16, 17, 18, 19]),
#                  titles=['Total Project Budget'])
#
#    check_parent_row(client=client,
#                  sheet=sheet,
#                  rowIdx=1,
#                  sumCols=set([6, 7, 12, 13, 15, 16, 17]),
#                  titles=['M&S Total','12.25% Overhead'])
#
#
#    # First walk through the rows and create a parent list
#    parents = set()
#    for rowIdx in range(2,len(sheet.rows)):
#        parents.add(sheet.rows[rowIdx].parent_id)
#
#    inMS = True
#
#    for rowIdx in range(2,len(sheet.rows)):
#
#        # Start of labor section
#        if sheet.rows[rowIdx].cells[0].value == 'Labor':
#            check_parent_row(client=client,
#                          sheet=sheet,
#                          rowIdx=rowIdx,
#                          sumCols=set([2, 6, 7, 12, 13, 14, 15, 16, 17, 18, 19]),
#                          titles=['Labor', 'SLAC Labor Rate FY21 (loaded): $274.63; Slac Tech Rate FY21 (loaded): $163.47'])
#
#            inMS = False
#
#        # This is a parent row
#        elif sheet.rows[rowIdx].id in parents:
#
#            # M&S Parent Rows
#            if inMS:
#                check_parent_row(client=client,
#                                 sheet=sheet,
#                                 rowIdx=rowIdx,
#                                 sumCols=set([6, 7, 12, 13, 15, 16, 17]),
#                                 titles=[])
#
#            # Labor Parent Rows
#            else:
#                check_parent_row(client=client,
#                                 sheet=sheet,
#                                 rowIdx=rowIdx,
#                                 sumCols=set([6, 7, 12, 13, 14, 15, 16, 17, 18, 19]),
#                                 titles=[])
#
#        else:
#            check_task_row(client=client,
#                           sheet=sheet,
#                           rowIdx=rowIdx,
#                           inMS=inMS)
#
#

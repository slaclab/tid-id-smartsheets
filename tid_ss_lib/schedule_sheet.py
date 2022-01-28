
import smartsheet

def check_parent_row(*, client, sheet, rowIdx):
    formulas = { 7: '=NETWORKDAYS([Baseline Start]@row, [Baseline Finish]@row)',
                12: '=SUM(CHILDREN())',
                13: '=SUM(CHILDREN())' }

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for i in range(len(row.cells)):
        if i in formulas:

            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != formulas[i]:
                print(f"   Invalid value in row {rowIdx+1} col {i+1} in schedule file. Expected '{formulas[i]}'. Got '{row.cells[i].formula}'. Fixing")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = formulas[i]
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_row(*, client, sheet, rowIdx):
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
                print(f"   Invalid value in row {rowIdx+1} col {i+1} in schedule file. Expected '{formulas[i]}'. Got '{row.cells[i].formula}'. Fixing")
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[i].id
                new_cell.formula = formulas[i]
                new_cell.strict = False
                new_row.cells.append(new_cell)

    if len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_task_links(*, client, sheet, rowIdx, laborRows, laborSheet):
    links = { 8: 2, 16: 20 }

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for k,v in links.items():
        rowIdTar = laborRows[rowIdx-1]['data'].id
        colIdTar = laborRows[rowIdx-1]['data'].cells[v].column_id
        shtIdTar = laborSheet.id

        if row.cells[k].link_in_from_cell is None or \
            row.cells[k].link_in_from_cell.row_id != rowIdTar or \
            row.cells[k].link_in_from_cell.column_id != colIdTar or \
            row.cells[k].link_in_from_cell.sheet_id != shtIdTar:

            print(f"   Incorrect schedule link for row {rowIdx+1} column {k+1}. Fixing")

            cell_link = smartsheet.models.CellLink()
            cell_link.sheet_id = shtIdTar
            cell_link.row_id = rowIdTar
            cell_link.column_id = colIdTar

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = row.cells[k].column_id
            new_cell.value = smartsheet.models.ExplicitNull()
            new_cell.link_in_from_cell = cell_link
            new_row.cells.append(new_cell)

    if len(new_row.cells) != 0:
        print(f"   Applying fixes to row {rowIdx+1}.")
        client.Sheets.update_rows(sheet.id, [new_row])


# Find stale or broken links to budget file
def check_broken(*, client, sheet):
    delList = []

    for rowIdx in range(1,len(sheet.rows)):
        row = sheet.rows[rowIdx]

        if row.cells[0].link_in_from_cell is None or row.cells[0].link_in_from_cell.status != 'OK':
            print(f"   Found stale row {rowIdx+1} in schedule file. Fixing")
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
            laborRows[i]['link'] = sheet.rows[rowIdx]
            rowIdx += 1
        else:

            print(f"   Adding new schedule row at position {rowIdx+1}, budget row = {laborRows[i]['data'].row_number}, parent = {laborRows[i]['parent']}")

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
            check_task_links(client=client, sheet=sheet, rowIdx=i+1, laborRows=laborRows, laborSheet=laborSheet)

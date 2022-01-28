
def check(*, row):

    if len(row.cells) != 21:
        raise Exception("Wrong number of columns in row in budget file")

    if row.cells[0].value != 'Total Project Budget':
        raise Exception("Incorrect row 0, col 0 value in budget file")

    sumChildrenCells = set([6, 7, 12, 13, 15, 16, 17, 18, 19])
    sumChildrenValue = '=SUM(CHILDREN())'

    for i in range(1,21):

        if i in sumChildrenCells:
            if not hasattr(row.cells[i],'formula') or row.cells[i].formula != sumChildrenValue:
                raise Exception (f"Invalid value in row 0 cell {i} in budget file. Expected '{sumChildrenValue}'. Got '{row.cells[i].formula}'")

        else:
            if row.cells[i].value is not None:
                raise Exception (f"Invalid value in row 0 cell {i} in budget file. Expected ''. Got '{row.cells[i].value}'")

#-----------------------------------------------------------------------------
# Title      : Manipulate Tracking Sheet
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

# Set formats
#
# https://smartsheet-platform.github.io/api-docs/#formatting
#
# Colors 31 = Dark Blue
#        26 = Dark Gray
#        23 = Blue
#        18 - Gray
#           = White

form = [ ",,,,,,,,,18,,,,,,,",  # 0
        ",,,,,,,,,,,,,,,,",  # 1
        ",,,,,,,,,18,,13,,,,,",  # 2
        ",,,,,,,,,18,,13,,,,,",  # 3
        ",,,,,,,,,18,,13,,,,,",  # 4
        ",,,,,,,,,18,,13,,,,,",  # 5
        ",,,,,,,,,18,,13,,,,,",  # 6
        ",,,,,,,,,,,13,,,,,",  # 7
        ",,,,,,,,,18,,13,,,,,",  # 8
        ",,,,,,,,,18,,13,,,,,",  # 9
        ",,,,,,,,,18,,13,,,,,",  # 10
        ",,,,,,,,,18,,13,,,,,",  # 11
        ",,,,,,,,,18,,13,,,,,",  # 12
        ",,,,,,,,,18,,13,,,,,",  # 13
        ",,,,,,,,,18,,,,,,,",  # 14
        ",,,,,,,,,18,,,,,,,",  # 15
        ",,,,,,,,,18,,,,,,,",  # 16
        ",,,,,,,,,,,,,,,,",  # 17
        ",,,,,,,,,,,,,,,," ]  # 18

Columns = ['Status Month',  # 0
           'Lookup PA',  # 1
           'Monthly Actuals Date',  # 2
           'Total Funds From Finance',  # 3
           'Monthly Actuals From Finance',  # 4
           'Total Actuals From Finance',  # 5
           'Remaining Funds From Finance',  # 6
           'Actuals Adjustment',  # 7
           'Reported Cost',  # 8
           'Budget Variance',  # 9
           'Budget Variance With Contingency',  # 10
           'Schedule Variance',  # 11
           'Schedule Variance With Contingency',  # 12
           'Reporting Variance',  # 13
           'Tracking Risk',  # 14
           'Budget Risk',  # 15
           'Schedule Risk',  # 16
           'Scope Risk',  # 17
           'Description Of Status']  # 18

RefName = 'Actuals Range 3'

def fix_structure(*, client, div, sheet):
    return False

    if div == 'id':
        labor_rate = navigate.TID_ID_RATE_NOTE
    elif div == 'cds':
        labor_rate = navigate.TID_CDS_RATE_NOTE

    if len(sheet.columns) != 19:
        print(f"   Wrong number of columns in tracking file, could not fix: Got {len(sheet.columns)}.")
        return False

    # Modify Column Names
    col11 = smartsheet.models.Column({'title': Columns[11],
                                      'type': 'TEXT_NUMBER',
                                      'index': 11})

    col12 = smartsheet.models.Column({'title': Columns[12],
                                      'type': 'TEXT_NUMBER',
                                      'index': 12})

    client.Sheets.update_column(sheet.id, sheet.columns[11].id, col11)
    client.Sheets.update_column(sheet.id, sheet.columns[12].id, col12)

    if div == 'id':
        xref = smartsheet.models.CrossSheetReference({
            'name': RefName,
            'source_sheet_id': navigate.TID_ID_ACTUALS_SHEET,
            'start_row_id': navigate.TID_ID_ACTUALS_START_ROW,
            'end_row_id': navigate.TID_ID_ACTUALS_END_ROW, })

    elif div == 'cds':
        xref = smartsheet.models.CrossSheetReference({
            'name': RefName,
            'source_sheet_id': navigate.TID_CDS_ACTUALS_SHEET,
            'start_row_id': navigate.TID_CDS_ACTUALS_START_ROW,
            'end_row_id': navigate.TID_CDS_ACTUALS_END_ROW, })

    client.Sheets.create_cross_sheet_reference(sheet.id, xref)
    return True

def check_structure(*, sheet):

    # Check column count
    if len(sheet.columns) != len(Columns):
        print(f"   Wrong number of columns in tracking file: Got {len(sheet.columns)} Expect {len(Columns)}.")
        return False

    else:
        ret = True

        for i,v in enumerate(Columns):
            if sheet.columns[i].title != v:
                print(f"   Mismatch tracking column name for col {i+1}. Got {sheet.columns[i].title}. Expect {v}.")
                ret = False

        return ret


def check_first_row(*, client, sheet, budgetSheet, doFixes):
    rowIdx = 0
    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    links = { 8: 15,
              9: 16,
             10: 17,
             11: 21,
             12: 22}

    noChange = set([0, 1, 7, 17, 18])

    formulas = {  2: '=VLOOKUP([Lookup PA]@row, {' + RefName + '}, 7, false)',
                  3: '=VLOOKUP([Lookup PA]@row, {' + RefName + '}, 6, false)',
                  4: '=VLOOKUP([Lookup PA]@row, {' + RefName + '}, 3, false)',
                  5: '=VLOOKUP([Lookup PA]@row, {' + RefName + '}, 4, false)',
                  6: '=VLOOKUP([Lookup PA]@row, {' + RefName + '}, 5, false)',
                 13: '=([Total Actuals From Finance]@row + [Actuals Adjustment]@row) - [Reported Cost]@row',
                 14: '=IF(ABS([Reporting Variance]@row) > 50000, "High", IF(ABS([Reporting Variance]@row) > 5000, "Medium", "Low"))',
                 15: '=IF([Budget Variance]@row > 50000, "High", IF([Budget Variance]@row > 5000, "Medium", "Low"))',
                 16: '=IF([Schedule Variance]@row > 50000, "High", IF([Schedule Variance]@row > 5000, "Medium", "Low"))' }

    for k,v in formulas.items():
        if ((not hasattr(row.cells[k],'formula')) or row.cells[k].formula != v) or row.cells[k].format != form[k]:

            if not hasattr(row.cells[k],'formula'):
                print(f"   Missing formula in row {rowIdx+1} cell {k+1} in tracking file. Expect {v}")

            elif row.cells[k].formula != v:
                print(f"   Incorrect formula in row {rowIdx+1} cell {k+1} in tracking file. Got {row.cells[k].formula} Expect {v}")

            if row.cells[k].format != form[k]:
                print(f"   Incorrect format in row {rowIdx+1} cell {k+1} in tracking file. Got '{row.cells[k].format}' Expect '{form[k]}'")

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[k].id
            new_cell.formula = v
            new_cell.format = form[k]
            new_cell.strict = False
            new_row.cells.append(new_cell)

    for k in noChange:
        if (row.cells[k].format != form[k]) and not (form[k] == ",,,,,,,,,,,,,,,," and row.cells[k].format is None):
            print(f"   Incorrect format in row {rowIdx+1} cell {k+1} in tracking file. Got '{row.cells[k].format}' Expect '{form[k]}'")
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[k].id

            if row.cells[k].value is None:
                new_cell.value = ''
            else:
                new_cell.value = row.cells[k].value
            new_cell.format = form[k]
            new_cell.strict = False
            new_row.cells.append(new_cell)

    # Need re-link
    relink = set()

    for k,v in links.items():
        if row.cells[k].format != form[k]:
            print(f"   Incorrect format in row {rowIdx+1} cell {k+1} in tracking file. Got '{row.cells[k].format}' Expect '{form[k]}'")
            relink.add(k)
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = row.cells[k].column_id
            new_cell.value = smartsheet.models.ExplicitNull()
            new_cell.format = form[k]
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to tracking row {row.row_number}.")
        client.Sheets.update_rows(sheet.id, [new_row])

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for k, v in links.items():
        rowIdTar = budgetSheet.rows[0].id
        colIdTar = budgetSheet.rows[0].cells[v].column_id
        shtIdTar = budgetSheet.id

        if k in relink or row.cells[k].link_in_from_cell is None or \
            row.cells[k].link_in_from_cell.row_id != rowIdTar or \
            row.cells[k].link_in_from_cell.column_id != colIdTar or \
            row.cells[k].link_in_from_cell.sheet_id != shtIdTar:

            print(f"   Incorrect tracking link for row {row.row_number} column {k+1}.")

            cell_link = smartsheet.models.CellLink()
            cell_link.sheet_id  = shtIdTar
            cell_link.row_id = rowIdTar
            cell_link.column_id = colIdTar

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = row.cells[k].column_id
            new_cell.value = smartsheet.models.ExplicitNull()
            new_cell.link_in_from_cell = cell_link
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to tracking row {row.row_number}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_other_row(*, client, rowIdx, sheet, doFixes):
    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for k in range(len(Columns)):

        if (hasattr(row.cells[k],'formula') and row.cells[k].formula is not None) or \
           (row.cells[k].link_in_from_cell is not None) or \
           ((row.cells[k].format != form[k]) and not (form[k] == ",,,,,,,,,,,,,,,," and row.cells[k].format is None)):

            if hasattr(row.cells[k],'formula') and row.cells[k].formula is not None:
                print(f"   Invalid formula in row {rowIdx+1} cell {k+1} in tracking file. Formula = {row.cells[k].formula}, Value = {row.cells[k].value}")

            if row.cells[k].link_in_from_cell is not None:
                print(f"   Invalid link in row {rowIdx+1} cell {k+1} in tracking file. Value = {row.cells[k].value}")

            if (row.cells[k].format != form[k]) and not (form[k] == ",,,,,,,,,,,,,,,," and row.cells[k].format is None):
                print(f"   Incorrect format in row {rowIdx+1} cell {k+1} in tracking file. Got '{row.cells[k].format}' Expect '{form[k]}'")

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
        print(f"   Applying fixes to tracking row {row.row_number}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check(*, client, sheet, budgetSheet, doFixes):
    check_first_row(client=client, sheet=sheet, budgetSheet=budgetSheet, doFixes=doFixes)

    for i in range(1,13):
        check_other_row(client=client, rowIdx=i, sheet=sheet, doFixes=doFixes)

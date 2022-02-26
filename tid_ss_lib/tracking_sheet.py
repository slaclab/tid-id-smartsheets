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

form = [ ",,,,,,,,,18,,,,,,,",   ",,,,,,,,,,,,,,,,",     ",,,,,,,,,18,,13,,,,,", ",,,,,,,,,18,,13,,,,,",
         ",,,,,,,,,18,,13,,,,,", ",,,,,,,,,18,,13,,,,,", ",,,,,,,,,18,,13,,,,,", ",,,,,,,,,18,,13,,,,,",
         ",,,,,,,,,18,,13,,,,,", ",,,,,,,,,18,,,,,,,",   ",,,,,,,,,18,,,,,,,",   ",,,,,,,,,18,,13,,,,,",
         ",,,,,,,,,18,,,,,,,",   ",,,,,,,,,18,,,,,,,",   ",,,,,,,,,18,,,,,,,",   ",,,,,,,,,,,,,,,,",
         ",,,,,,,,,,,,,,,," ]

def check_structure(*, sheet):

    columns = ['Status Month',
               'Lookup PA',
               'Monthly Actuals From Finance',
               'Total Actuals From Finance',
               'Funds Remaining From Finance',
               'Actuals Adjustment',
               'Reported Cost',
               'Budget Variance',
               'Budget Variance With Contingency',
               'Duration Variance',
               'Duration Variance With Contingency',
               'Reporting Variance',
               'Tracking Risk',
               'Budget Risk',
               'Schedule Risk',
               'Scope Risk',
               'Description Of Status']

    # Check column count
    if len(sheet.columns) != len(columns):
        print(f"   Wrong number of columns in tracking file: Got {len(sheet.columns)} Expect {len(columns)}.")
        return False

    else:
        ret = True

        for i,v in enumerate(columns):
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

    links = { 6: 15,
              7: 16,
              8: 17,
              9: 18,
             10: 19}

    noChange = set([0, 1, 5, 15, 16])

    formulas = {  2: '=VLOOKUP([Lookup PA]@row, {Actuals Range 1}, 3, false)',
                  3: '=VLOOKUP([Lookup PA]@row, {Actuals Range 1}, 4, false)',
                  4: '=VLOOKUP([Lookup PA]@row, {Actuals Range 1}, 5, false)',
                 11: '=([Total Actuals From Finance]@row + [Actuals Adjustment]@row) - [Reported Cost]@row',
                 12: '=IF([Reporting Variance]@row > 50000, "High", IF([Reporting Variance]@row > 5000, "Medium", "Low"))',
                 13: '=IF([Budget Variance]@row > 50000, "High", IF([Budget Variance]@row > 5000, "Medium", "Low"))',
                 14: '=IF([Duration Variance With Contingency]@row > 300, "High", IF([Duration Variance With Contingency]@row > 100, "Medium", "Low"))' }

    for k,v in formulas.items():
        if ((not hasattr(row.cells[k],'formula')) or row.cells[k].formula != v) or row.cells[k].format != form[k]:

            if ((not hasattr(row.cells[k],'formula')) or row.cells[k].formula != v):
                print(f"   Incorrect formula in row {rowIdx+1} cell {k+1} in tracking file.")

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

    for k in range(len(row.cells)):
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

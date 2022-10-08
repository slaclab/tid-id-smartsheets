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
import re
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

form = [ ",,,,,,,,,18,,,,,,,",      # 0
        ",,,,,,,,,,,,,,,,",         # 1
        ",,,,,,,,,18,,13,,,,,",     # 2
        ",,,,,,,,,18,,13,2,1,2,,",  # 3
        ",,,,,,,,,18,,13,2,1,2,,",  # 4
        ",,,,,,,,,18,,13,2,1,2,,",  # 5
        ",,,,,,,,,18,,13,2,1,2,,",  # 6
        ",,,,,,,,,18,,13,2,1,2,,",  # 7
        ",,,,,,,,,18,,13,2,1,2,,",  # 8
        ",,,,,,,,,18,,13,2,1,2,,",  # 9
        ",,,,,,,,,18,,13,2,1,2,,",  # 10
        ",,,,,,,,,18,,13,2,1,2,,",  # 11
        ",,,,,,,,,18,,,,,,,",       # 12
        ",,,,,,,,,18,,,,,,,",       # 13
        ",,,,,,,,,18,,,,,,,",       # 14
        ",,,,,,,,,,,,,,,,",         # 15
        ",,,,,,,,,,,,,,,," ]        # 16

Columns = ['Status Month',  # 0
           'Lookup PA',  # 1
           'Monthly Actuals Date',  # 2
           'Monthly Actuals From Finance',  # 3
           'Total Actuals From Finance',  # 4
           'Total Budget', # 5
           'Actual Cost',  #  6
           'Remaining Funds', # 7
           'Cost Variance',  #  8
           'Cost Variance With Contingency', # 9
           'Schedule Variance', # 10
           'Reporting Variance', # 11
           'Tracking Risk',  # 12
           'Budget Risk',  # 13
           'Schedule Risk',  # 14
           'Scope Risk',  # 15
           'Description Of Status']  # 16

RefName = None

def fix_structure(*, client, div, sheet):
    global refName

    if len(sheet.columns) != 16:
        print(f"   Can't fix tracking sheet. Got {len(sheet.columns)} Expect {16}.")
        return False

    print("   Fixing tracking sheet.")

    col5 = smartsheet.models.Column({'title': Columns[5],
                                     'type': 'TEXT_NUMBER',
                                     'index': 5})

    client.Sheets.add_columns(sheet.id, [col5])

    return False

    RefName = 'Actuals Range 3'

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


def check_first_row(*, client, sheet, projectSheet, doFixes, cData):
    global RefName

    rowIdx = 0
    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    links = { 5: cData['Total Budgeted Cost']['position'],
              6: cData['Actual Cost']['position'],
              7: cData['Remaining Funds']['position'],
              8: cData['Cost Variance']['position'],
              9: cData['Cost Variance With Contingency']['position'],
              10: cData['Schedule Variance']['position'] }

    noChange = set([0, 1, 15, 16])

    if RefName is None:
        if hasattr(row.cells[2],'formula'):
            RefStr = re.findall(r"\{.*?\}",row.cells[2].formula)[0];
        else:
            RefStr = '{Actuals Range 1}'
    else:
        RefStr = f"{Refname}"

    #print(f"Using RefStr = {RefStr}")

    formulas = {  2: '=VLOOKUP([Lookup PA]@row, ' + RefStr + ', 5, false)',
                  3: '=VLOOKUP([Lookup PA]@row, ' + RefStr + ', 3, false)',
                  4: '=VLOOKUP([Lookup PA]@row, ' + RefStr + ', 4, false)',
                 11: '=([Actual Cost]@row - [Total Actuals From Finance]@row)',
                 12: '=IF(ABS([Reporting Variance]@row) < -50000, "High", IF(ABS([Reporting Variance]@row) < -5000, "Medium", "Low"))',
                 13: '=IF([Cost Variance]@row < -50000, "High", IF([Cost Variance]@row < -5000, "Medium", "Low"))',
                 14: '=IF([Schedule Variance]@row < -50000, "High", IF([Schedule Variance]@row < -5000, "Medium", "Low"))' }

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
        rowIdTar = projectSheet.rows[0].id
        colIdTar = projectSheet.rows[0].cells[v].column_id
        shtIdTar = projectSheet.id

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


def check(*, client, sheet, projectSheet, doFixes, div, cData):
    if not check_structure(sheet=sheet):
        #fix_structure(client=client, div=div, sheet=sheet)
        return False

    check_first_row(client=client, sheet=sheet, projectSheet=projectSheet, doFixes=doFixes, cData=cData)

    for i in range(1,13):
        check_other_row(client=client, rowIdx=i, sheet=sheet, doFixes=doFixes)


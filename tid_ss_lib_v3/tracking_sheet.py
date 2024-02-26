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
from . import tracking_sheet_columns


def find_columns(*, client, sheet, doFixes, tData):

    # Look for columns we want to delete
    for tod in tracking_sheet_columns.ToDelete:
        for i in range(len(sheet.columns)):
            if sheet.columns[i].title == tod:
                if doFixes:
                    print(f"Found tracking column to delete {tod} at position {i+1}. Deleting")
                    client.Sheets.delete_column(sheet.id, sheet.columns[i].id)
                    return False

    # Look for each expected column
    for k,v in tData.items():
        found = False
        for i in range(len(sheet.columns)):
            if sheet.columns[i].title == k:
                found = True
                if v['position'] != i:
                    print(f"Tracking column location mismatch for {k}. Expected at {v['position']+1}, found at {i+1}.")
                    v['position'] = i

        if not found:
            print(f"Tracking column not found: {k}.")

            if doFixes is True:
                print(f"Adding tracking column: {k}.")
                col = smartsheet.models.Column({'title': k,
                                                'type': v['type'],
                                                'index': v['position']})

                client.Sheets.add_columns(sheet.id, [col])

            return False

    return True


def check_first_row(*, client, sheet, doFixes, tData, projectSheet, actualsSheet, cData):
    row = sheet.rows[0]

    relink = []

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for k,v in tData.items():
        doFormat = False
        idx = v['position']

        # First check format
        if v['format'] is not None and row.cells[idx].format != v['format']:
            print(f"   Incorrect format in row 0 cell {idx+1} in tracking file. Got {row.cells[idx].format} Exp {v['format']}")
            doFormat = True

        # Row has a formula
        if v['formula'] is not None:
            doRow = False

            if not hasattr(row.cells[idx],'formula') or row.cells[idx].formula != v['formula']:
                print(f"   Incorrect formula in row 0 cell {idx+1} in tracking file. Got {row.cells[idx].formula} Exp {v['formula']}")
                doRow = True

            if doRow or doFormat:
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[idx].id
                new_cell.formula = v['formula']

                if v['format'] is not None:
                    new_cell.format = v['format']

                new_cell.strict = False
                new_row.cells.append(new_cell)

        # Row has a forced value
        elif v['forced'] is not None:
            doRow = False

            if (not (row.cells[idx].value is None and v['forced'] == '')) and (row.cells[idx].value is None or row.cells[idx].value != v['forced']):
                print(f"   Incorrect value in row 0 cell {idx+1} in tracking file. Got {row.cells[idx].value} Exp {v['forced']}")
                doRow = True

            if doRow or doFormat:
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[idx].id
                new_cell.value = v['forced']

                if v['format'] is not None:
                    new_cell.format = v['format']

                new_cell.strict = False
                new_row.cells.append(new_cell)

        # Catch format only case
        elif doFormat:

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[idx].id

            if v['link'] is not None:
                new_cell.value = smartsheet.models.ExplicitNull()
                relink.append(k)
            elif row.cells[idx].value is None:
                new_cell.value = ''
            else:
                new_cell.value = row.cells[idx].value

            new_cell.format = v['format']
            new_cell.strict = False
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print("   Applying fixes to tracking row 0")
        client.Sheets.update_rows(sheet.id, [new_row])

    # Last check links

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for k,v in tData.items():
        idx = v['position']

        if v['link'] is not None:

            if v['link'][0] == 'project':
                remCol = cData[v['link'][1]]['position']

                rowIdTar = projectSheet.rows[0].id
                colIdTar = projectSheet.rows[0].cells[remCol].column_id
                shtIdTar = projectSheet.id
            else:
                remCol = v['link'][1]

                rowIdTar = actualsSheet.rows[0].id
                colIdTar = actualsSheet.rows[0].cells[remCol].column_id
                shtIdTar = actualsSheet.id

            if k in relink or row.cells[idx].link_in_from_cell is None or \
                row.cells[idx].link_in_from_cell.row_id != rowIdTar or \
                row.cells[idx].link_in_from_cell.column_id != colIdTar or \
                row.cells[idx].link_in_from_cell.sheet_id != shtIdTar:

                print(f"   Incorrect tracking link for row {row.row_number} column {idx+1}.")

                cell_link = smartsheet.models.CellLink()
                cell_link.sheet_id  = shtIdTar
                cell_link.row_id = rowIdTar
                cell_link.column_id = colIdTar

                new_cell = smartsheet.models.Cell()
                new_cell.column_id = row.cells[idx].column_id
                new_cell.value = smartsheet.models.ExplicitNull()
                new_cell.link_in_from_cell = cell_link
                new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to tracking row {row.row_number}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check_other_row(*, client, rowIdx, sheet, tData, doFixes):
    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    if row.cells[0].value is None or row.cells[0].value == "":
        return

    for k,v in tData.items():
        idx = v['position']

        if (hasattr(row.cells[idx],'formula') and row.cells[idx].formula is not None) or \
           (row.cells[idx].link_in_from_cell is not None) or \
           ((row.cells[idx].format != v['format']) and not (v['format'] == ",,,,,,,,,,,,,,,," and row.cells[idx].format is None)):

            if hasattr(row.cells[idx],'formula') and row.cells[idx].formula is not None:
                print(f"   Invalid formula in row {rowIdx+1} cell {idx+1} in tracking file. Formula = {row.cells[idx].formula}, Value = {row.cells[idx].value}")

            if row.cells[idx].link_in_from_cell is not None:
                print(f"   Invalid link in row {rowIdx+1} cell {idx+1} in tracking file. Value = {row.cells[idx].value}")

            if (row.cells[idx].format != v['format']) and not (v['format'] == ",,,,,,,,,,,,,,,," and row.cells[idx].format is None):
                print(f"   Incorrect format in row {rowIdx+1} cell {idx+1} in tracking file. Got '{row.cells[idx].format}' Expect {v['format']}")

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[idx].id

            if row.cells[idx].value is None:
                new_cell.value = ''
            else:
                new_cell.value = row.cells[idx].value

            new_cell.format = v['format']
            new_cell.strict = False
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to tracking row {row.row_number}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check(*, client, sheet, projectSheet, actualsSheet, tData, doFixes, cData, doDownload, div):

    # First check structure
    while True:
        ret = find_columns(client=client, sheet=sheet, doFixes=doFixes, tData=tData )

        if ret is False and doFixes is False:
            return False

        elif ret is True:
            break

        else:
            sheet = client.Sheets.get_sheet(sheet.id, include='format')

    check_first_row(client=client, sheet=sheet, projectSheet=projectSheet, actualsSheet=actualsSheet, doFixes=doFixes, cData=cData, tData=tData)

    for i in range(1, len(sheet.rows)):
        check_other_row(client=client, rowIdx=i, sheet=sheet, tData=tData, doFixes=doFixes)

    if isinstance(doDownload,str):
        client.Sheets.get_sheet_as_excel(sheet.id, doDownload)


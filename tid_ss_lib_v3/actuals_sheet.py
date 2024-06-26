#-----------------------------------------------------------------------------
# Title      : Actuals Sheet Manipulation
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

import datetime

StdColumns = 5
CostFormat = ",,,,,,,,,,,13,2,1,2,,"

def check(*, client, sheet, doFixes, resources):
    inLabor = False

    resFound = {v : False for v in resources}
    empty = []
    addRows = []
    parentId = None

    # Process the rows
    for rowIdx in range(0,len(sheet.rows)):

        if sheet.rows[rowIdx].cells[0].value == 'Labor Actuals':
            inLabor = True
            parentId = sheet.rows[rowIdx].id

        elif inLabor:
            resName = sheet.rows[rowIdx].cells[1].value

            if isinstance(resName, str) and resName != "":
                resFound[resName] = True
            else:
                empty.append(rowIdx)

    # Add new resource rows
    for k,v in resFound.items():

        if v is False:
            if doFixes:
                print(f"   Adding missing resource to actuals: {k}")
                new_row = smartsheet.models.Row()
                new_row.parent_id = parentId

                # Item
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[0].id
                new_cell.value = k
                new_cell.strict = False
                new_row.cells.append(new_cell)

                # Assigned To
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[1].id
                new_cell.value = k
                new_cell.strict = False
                new_row.cells.append(new_cell)

                # Actual Hours
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[2].id
                new_cell.value = "0"
                new_cell.strict = False
                new_row.cells.append(new_cell)

                # Actual Cost
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[3].id
                new_cell.value = "0"
                new_cell.format = CostFormat
                new_cell.strict = False
                new_row.cells.append(new_cell)

                # Current Cost
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[4].id
                new_cell.value = "0"
                new_cell.format = CostFormat
                new_cell.strict = False
                new_row.cells.append(new_cell)

                addRows.append(new_row)

            else:
                print(f"   Missing resource in actuals: {k}")

    # Process add rows
    if len(addRows) > 0:
        client.Sheets.add_rows(sheet.id, addRows)


def update_actuals_row(*, sheet, row, parentId, resName, data, months):

    new_row = smartsheet.models.Row()
    new_row.parent_id = parentId

    if row is not None:
        new_row.id = row.id

    # Item
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = sheet.columns[0].id
    new_cell.value = resName
    new_cell.strict = False
    new_row.cells.append(new_cell)

    # Assigned To
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = sheet.columns[1].id
    new_cell.value = resName
    new_cell.strict = False
    new_row.cells.append(new_cell)

    # Actual Hours
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = sheet.columns[2].id
    if data is not None:
        new_cell.value = data['total_hours']
    else:
        new_cell.value = 0.0
    new_cell.strict = False
    new_row.cells.append(new_cell)

    # Actual Cost
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = sheet.columns[3].id
    if data is not None:
        new_cell.value = data['total_cost']
    else:
        new_cell.value = 0.0
    new_cell.strict = False
    new_cell.format = CostFormat
    new_row.cells.append(new_cell)

    # Current Cost
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = sheet.columns[4].id
    if data is not None:
        new_cell.value = data['current_cost']
    else:
        new_cell.value = 0.0
    new_cell.strict = False
    new_cell.format = CostFormat
    new_row.cells.append(new_cell)

    for i in range(len(months)):
        new_cell = smartsheet.models.Cell()
        new_cell.column_id = sheet.columns[StdColumns+i].id

        if data is None or months[i] not in data['monthly_hours']:
            new_cell.value = 0
        else:
            new_cell.value = data['monthly_hours'][months[i]]

        new_cell.strict = False
        new_row.cells.append(new_cell)

    return new_row


def update_pas_row(*, sheet, row, parentId, paName, data, months):

    new_row = smartsheet.models.Row()
    new_row.parent_id = parentId

    if row is not None:
        new_row.id = row.id

    # Item
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = sheet.columns[0].id
    new_cell.value = paName
    new_cell.strict = False
    new_row.cells.append(new_cell)

    # Assigned To
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = sheet.columns[1].id
    new_cell.value = ''
    new_cell.strict = False
    new_row.cells.append(new_cell)

    # Actual Hours
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = sheet.columns[2].id
    new_cell.value = ''
    new_cell.strict = False
    new_row.cells.append(new_cell)

    # Actual Cost
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = sheet.columns[3].id
    if data is not None:
        new_cell.value = data['Total Actuals']
    else:
        new_cell.value = 0.0
    new_cell.strict = False
    new_cell.format = CostFormat
    new_row.cells.append(new_cell)

    # Current Cost
    new_cell = smartsheet.models.Cell()
    new_cell.column_id = sheet.columns[4].id
    if data is not None:
        new_cell.value = data['Monthly Actuals']
    else:
        new_cell.value = 0.0
    new_cell.strict = False
    new_cell.format = CostFormat
    new_row.cells.append(new_cell)

    for i in range(StdColumns,len(months)+StdColumns):
        new_cell = smartsheet.models.Cell()
        new_cell.column_id = sheet.columns[i].id
        new_cell.value = ''
        new_cell.strict = False
        new_row.cells.append(new_cell)

    return new_row


def update_columns(*, client, sheet, wbsData):
    refresh = False

    # Create found table
    colMap = {f'{v} Hours':None for v in wbsData['months']}

    # Look through existing columns
    for i in range(StdColumns, len(sheet.columns)):
        if sheet.columns[i].title in colMap:
            colMap[sheet.columns[i].title] = i
        else:
            #print(f"   Deleting column {sheet.columns[i].title}")
            client.Sheets.delete_column(sheet.id, sheet.columns[i].id)
            refresh = True

    # Add missing columns
    for i,v in enumerate(colMap):
        if colMap[v] is None:

            #print(f"   Adding Missing Column: {v} at {i+StdColumns}")
            newCol = smartsheet.models.Column({'title': v,
                                               'type': 'TEXT_NUMBER',
                                               'index': i + StdColumns})
            client.Sheets.add_columns(sheet.id, [newCol])
            refresh = True

    return refresh

def update_actuals_ms(*, client, sheet, wbsData):
    parentId = None

    if 'M&S' not in wbsData['person']:
        return

    # Process the rows
    for rowIdx in range(0,len(sheet.rows)):

        if sheet.rows[rowIdx].cells[0].value == 'M&S Actuals':
            parentId = sheet.rows[rowIdx].id

        elif parentId is not None and sheet.rows[rowIdx].cells[0].value == 'M&S':
            updRow = update_actuals_row(sheet=sheet, row=sheet.rows[rowIdx], parentId=parentId,
                                        resName='M&S', data=wbsData['person']['M&S'], months=wbsData['months'])
            client.Sheets.update_rows(sheet.id, [updRow])
            parentId = None


def update_actuals_labor(*, client, sheet, wbsData):
    addRows = []
    updRows = []
    parentId = None

    resFound = {v : False for v in wbsData['person'] if v != 'M&S'}

    # Process the rows
    for rowIdx in range(0,len(sheet.rows)):

        if sheet.rows[rowIdx].cells[0].value == 'Labor Actuals':
            parentId = sheet.rows[rowIdx].id

        elif parentId is not None:
            resName = sheet.rows[rowIdx].cells[1].value
            data = None

            if isinstance(resName, str) and resName != "":
                if resName in resFound and not resFound[resName]:
                    resFound[resName] = True
                    data = wbsData['person'][resName]

                updRows.append(update_actuals_row(sheet=sheet, row=sheet.rows[rowIdx], parentId=parentId, resName=resName, data=data, months=wbsData['months']))

    # Process update rows
    if len(updRows) > 0:
        client.Sheets.update_rows(sheet.id, updRows)

    # Add new resource rows
    for k,v in resFound.items():
        if v is False:
            addRows.append(update_actuals_row(sheet=sheet, row=None, parentId=parentId, resName=k, data=wbsData['person'][k], months=wbsData['months']))

    # Process add rows
    if len(addRows) > 0:
        client.Sheets.add_rows(sheet.id, addRows)


def update_actuals_pas(*, client, sheet, wbsData):
    addRows = []
    updRows = []
    parentId = None
    inPas = False

    paFound = {v : False for v in wbsData['pas']}

    # Process the rows
    for rowIdx in range(0,len(sheet.rows)):

        if sheet.rows[rowIdx].cells[0].value == 'Actuals Breakdown':
            inPas = False

        elif sheet.rows[rowIdx].cells[0].value == 'PA Actuals':
            parentId = sheet.rows[rowIdx].id
            inPas = True

        elif inPas:
            paName = sheet.rows[rowIdx].cells[0].value
            data = None

            if isinstance(paName, str) and paName != "":
                if paName in paFound and not paFound[paName]:
                    paFound[paName] = True
                    data = wbsData['pas'][paName]

                updRows.append(update_pas_row(sheet=sheet, row=sheet.rows[rowIdx], parentId=parentId, paName=paName, data=data, months=wbsData['months']))

    # Process update rows
    if len(updRows) > 0:
        client.Sheets.update_rows(sheet.id, updRows)

    # Add new resource rows
    for k,v in paFound.items():
        if v is False:
            addRows.append(update_pas_row(sheet=sheet, row=None, parentId=parentId, paName=k, data=wbsData['pas'][k], months=wbsData['months']))

    # Process add rows
    if len(addRows) > 0:
        client.Sheets.add_rows(sheet.id, addRows)


def update_actuals(*, client, sheet, wbsData):

    if update_columns(client=client, sheet=sheet, wbsData=wbsData):
        sheet = client.Sheets.get_sheet(sheet.id, include='format')

    update_actuals_pas(client=client, sheet=sheet, wbsData=wbsData)
    update_actuals_ms(client=client, sheet=sheet, wbsData=wbsData)
    update_actuals_labor(client=client, sheet=sheet, wbsData=wbsData)


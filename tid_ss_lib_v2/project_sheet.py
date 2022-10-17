#-----------------------------------------------------------------------------
# Title      : Schedule Sheet Manipulation
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
from . import project_sheet_columns

import datetime


def find_columns(*, client, sheet, doFixes, cData):

    for k,v in cData.items():
        found = False
        for i in range(len(sheet.columns)):
            if sheet.columns[i].title == k:
                found = True
                if v['position'] != i:
                    print(f"Column location mismatch for {k}. Expected at {v['position']}, found at {i}.")
                    v['position'] = i

                #print(f"Found column {k} at position {v['position']}, with id {sheet.columns[i].id}")

        if not found:
            print(f"Column not found: {k}.")

            if doFixes is True:
                print(f"Adding column: {k}.")
                col = smartsheet.models.Column({'title': k,
                                                'type': v['type'],
                                                'index': v['position']})

                client.Sheets.add_columns(sheet.id, [col])

            return False

    return True


def check_row(*, client, sheet, rowIdx, key, div, cData, doFixes):

    if div == 'id':
        laborRate = navigate.TID_ID_RATE_NOTE
    elif div == 'cds':
        laborRate = navigate.TID_CDS_RATE_NOTE

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for k,v in cData.items():
        doFormat = False
        idx = v['position']
        data = v[key]

        # First check format
        if data['format'] is not None and row.cells[idx].format != data['format']:
            print(f"   Incorrect format in row {rowIdx+1} {key} cell {idx+1} in project file. Got {row.cells[idx].format} Exp {data['format']}")
            doFormat = True

        # Row has a formula
        if data['formula'] is not None:
            doRow = False

            if not hasattr(row.cells[idx],'formula') or row.cells[idx].formula != data['formula']:
                print(f"   Incorrect formula in row {rowIdx+1} {key} cell {idx+1} in project file. Got {row.cells[idx].formula} Exp {data['formula']}")
                doRow = True

            if doRow or doFormat:
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[idx].id
                new_cell.formula = data['formula']

                if data['format'] is not None:
                    new_cell.format = data['format']

                new_cell.strict = False
                new_row.cells.append(new_cell)

        # Row has a forced value
        elif data['forced'] is not None:
            doRow = False

            if data['forced'] == 'MS_OVERHEAD':
                data['forced'] = navigate.OVERHEAD_NOTE
            elif data['forced'] == 'LAB_RATE':
                data['forced'] = laborRate

            if (not (row.cells[idx].value is None and data['forced'] == '')) and (row.cells[idx].value is None or row.cells[idx].value != data['forced']):
                print(f"   Incorrect value in row {rowIdx+1} {key} cell {idx+1} in project file. Got {row.cells[idx].value} Exp {data['forced']}")
                doRow = True

            if doRow or doFormat:
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[idx].id
                new_cell.value = data['forced']

                if data['format'] is not None:
                    new_cell.format = data['format']

                new_cell.strict = False
                new_row.cells.append(new_cell)

        # Row has a default value and is empty
        elif data['default'] is not None and (row.cells[idx].value is None or row.cells[idx].value == ''):
            print(f"   Missing default value in row {rowIdx+1} {key} cell {idx+1} in project file.")

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[idx].id
            new_cell.value = data['default']

            if data['format'] is not None:
                new_cell.format = data['format']

            new_cell.strict = False
            new_row.cells.append(new_cell)

        # Catch format only case
        elif doFormat:

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[idx].id

            if row.cells[idx].value is None:
                new_cell.value = ''
            else:
                new_cell.value = row.cells[idx].value
            new_cell.format = data['format']
            new_cell.strict = False
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to project row {rowIdx+1} {key}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def cost_ms(*, sheet, rowIdx, cData, msTable):

    # Extract key columns
    qty = float(sheet.rows[rowIdx].cells[cData['Budgeted Quantity']['position']].value)
    perItem = float(sheet.rows[rowIdx].cells[cData['Cost Per Item']['position']].value)
    endStr = sheet.rows[rowIdx].cells[cData['End']['position']].value

    if endStr is None:
        return

    # Split date fields
    endFields = endStr.split('T')[0].split('-')    # End date fields

    # Convert to python date object
    endDate = datetime.date(int(endFields[0]),int(endFields[1]),int(endFields[2]))

    # Generate a key for the end year/month
    k = f"{endDate.year}_{endDate.month:02d}"

    # Add year/month to dictionary if it does not exist
    if k not in msTable:
        msTable[k] = 0.0

    # Increment the ms cost for that year/month
    msTable[k] += (qty * perItem)


def cost_labor(*, sheet, rowIdx, cData, laborTable):

    # Extract key columns
    hours = float(sheet.rows[rowIdx].cells[cData['Budgeted Quantity']['position']].value)
    rate = sheet.rows[rowIdx].cells[cData['Cost Per Item']['position']].value
    startStr = sheet.rows[rowIdx].cells[cData['Start']['position']].value
    endStr = sheet.rows[rowIdx].cells[cData['End']['position']].value

    if endStr is None or startStr is None:
        return

    # Split date fields
    startFields = startStr.split('T')[0].split('-')  # Start date fields
    endFields = endStr.split('T')[0].split('-')    # End date fields

    # Convert to python date object
    startDate = datetime.date(int(startFields[0]),int(startFields[1]),int(startFields[2]))
    endDate = datetime.date(int(endFields[0]),int(endFields[1]),int(endFields[2]))

    # Count the number of weekdays in the time period
    days = 0

    for n in range(int((endDate - startDate).days)+1):
        dt = startDate + datetime.timedelta(n)

        if dt.weekday() != 5 and dt.weekday() != 6:
            days += 1

    # Compute hours per day
    hoursPerDay = hours / days

    # Iterate through the weekdays
    for n in range(int((endDate - startDate).days)+1):
        dt = startDate + datetime.timedelta(n)

        if dt.weekday() != 5 and dt.weekday() != 6:

            # Generate a key for the year/month
            k = f"{dt.year}_{dt.month:02d}"

            # Add year/month to dictionary if it does not exist
            if k not in laborTable:
                laborTable[k] = {rate: 0.0}

            # Add rate to sub-dictionrary for month
            elif rate not in laborTable[k]:
                laborTable[k][rate] = 0.0

            # Increment the number of hours in that year/month, for the given rate
            laborTable[k][rate] += hoursPerDay


def check(*, client, sheet, doFixes, div, cData, doCost, name, doDownload):
    inLabor = False
    inMS = False
    msTable = {}
    laborTable = {}

    # First check structure
    while True:
        ret = find_columns(client=client, sheet=sheet, doFixes=doFixes, cData=cData)

        if ret is False and doFixes is False:
            return False

        elif ret is True:
            break

        else:
            sheet = client.Sheets.get_sheet(sheet.id, include='format')

    # First do row 0
    check_row(client=client, sheet=sheet, rowIdx=0, key='top', div=div, cData=cData, doFixes=doFixes)

    # First walk through the rows and create a parent list
    parents = set()
    for rowIdx in range(len(sheet.rows)):
        parents.add(sheet.rows[rowIdx].parent_id)

    # Process the rest of the rows
    for rowIdx in range(1,len(sheet.rows)):

        # MS Section
        if sheet.rows[rowIdx].cells[0].value == 'M&S':
            check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='ms_top', div=div, cData=cData, doFixes=doFixes)
            inMS = True
            inLabor = False

        elif sheet.rows[rowIdx].cells[0].value == 'Labor':
            check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='labor_top', div=div, cData=cData, doFixes=doFixes)
            inMS = False
            inLabor = True

        elif inMS:

            # In parent
            if sheet.rows[rowIdx].id in parents:
                check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='ms_parent', div=div, cData=cData, doFixes=doFixes)
            else:
                check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='ms_task', div=div, cData=cData, doFixes=doFixes)
                if doCost:
                    cost_ms(sheet=sheet, rowIdx=rowIdx, cData=cData, msTable=msTable)

        elif inLabor:

            # In parent
            if sheet.rows[rowIdx].id in parents:
                check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='labor_parent', div=div, cData=cData, doFixes=doFixes)
            else:
                check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='labor_task', div=div, cData=cData, doFixes=doFixes)
                if doCost:
                    cost_labor(sheet=sheet, rowIdx=rowIdx, cData=cData, laborTable=laborTable)


    # Generate excel friend view of monthly spending
    if doCost:
        write_cost_table(name=name, laborTable=laborTable, msTable=msTable)

    if doDownload is not False:
        client.Sheets.get_sheet_as_excel(sheet.id, doDownload)

    return True


def write_cost_table(*, name, laborTable, msTable):

    # Generate excel friend view of monthly spending
    months = set([])
    rates = set([])

    # First generate month and rate lists to create excel grid
    for mnth in laborTable:
        months.add(mnth)
        for rte in laborTable[mnth]:
            rates.add(rte)

    for mnth in msTable:
        months.add(mnth)

    months = sorted(months)
    rates = sorted(rates)

    with open(f'{name} Cost.csv', 'w') as f:

        # Create Header
        f.write('Month,M&S,')

        for rte in rates:
            f.write(f'{float(rte):0.2f} Hours, {float(rte):0.2f} Dollars,')

        f.write("Total Labor Hours, Total Labor Dollars, Total Dollars\n")

        # Each month
        for m in months:
            totLabHours = 0.0
            totLabDollars = 0.0
            totDollars = 0.0

            # Put in M&S
            if m in msTable:
                val = float(msTable[m])
            else:
                val = 0.0

            f.write(f'{m}, {val:0.2f},')
            totDollars += val

            # Put in Labor
            for rte in rates:
                if m in laborTable and rte in laborTable[m]:
                    hrs = float(laborTable[m][rte])
                    val = float(hrs) * float(rte)
                else:
                    hrs = 0.0
                    val = 0.0

                totLabHours += hrs
                totLabDollars += val
                totDollars += val

                f.write(f'{hrs:0.2f}, {val:0.2f},')

            f.write(f"{totLabHours:0.2f}, {totLabDollars:0.2f},{totDollars:0.2f}\n")


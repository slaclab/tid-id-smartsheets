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
from . import configuration

import datetime


def find_columns(*, client, sheet, doFixes, cData):

    # Look for columns we want to delete
    for tod in project_sheet_columns.ToDelete:
        for i in range(len(sheet.columns)):
            if sheet.columns[i].title == tod:
                if doFixes:
                    print(f"   Found project column to delete {tod} at position {i+1}. Deleting")
                    client.Sheets.delete_column(sheet.id, sheet.columns[i].id)
                    return False

    # Look for each expected column
    for k,v in cData.items():
        found = False
        for i in range(len(sheet.columns)):
            if sheet.columns[i].title == k:
                found = True
                if v['position'] != i:
                    print(f"   Project column location mismatch for {k}. Expected at {v['position']+1}, found at {i+1}.")
                    v['position'] = i

        if not found:
            print(f"   Project column not found: {k}.")

            if doFixes is True:
                print(f"   Adding project column: {k}.")
                col = smartsheet.models.Column({'title': k,
                                                'type': v['type'],
                                                'index': v['position']})

                client.Sheets.add_columns(sheet.id, [col])

            return False

    return True

def check_row(*, client, sheet, rowIdx, key, div, cData, doFixes, doTask, resources):

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    # Skip empty rows
    if row.cells[0].value is None or row.cells[0].value == "":
        #print(f"   Skipping empty row {rowIdx+1}")
        return

    # Detect labor milestones
    if key == 'labor_task':
        notesIdx = cData['Item Notes']['position']

        if row.cells[notesIdx].value is not None and isinstance(row.cells[notesIdx].value,str) and "Milestone" in row.cells[notesIdx].value:
            key = 'labor_mstone'

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
                data['forced'] = div.ms_overhead
            elif data['forced'] == 'LAB_RATE':
                data['forced'] = div.rate_note

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

    # Extract resource list
    if resources is not None:
        res = row.cells[cData['Assigned To']['position']].value
        if isinstance(res,str):
            resources.add(row.cells[cData['Assigned To']['position']].value)

    # Check for tasks with non zero hours and not assigned
    if doTask and key == 'labor_task':
        tsk = row.cells[cData['Task']['position']].value
        qty = row.cells[cData['Budgeted Quantity']['position']].value
        asg = row.cells[cData['Assigned To']['position']].value
        dur = row.cells[cData['Duration']['position']].value

        if dur is not None and 'd' in dur:
            days = int(dur[:-1])

            if days > 60:
                print(f"   Duration Warning: Task {tsk} on row {rowIdx+1} has duration {days} days")

        elif dur is not None and 'w' in dur:
            weeks = int(dur[:-1])

            if weeks > 12:
                print(f"   Duration Warning: Task {tsk} on row {rowIdx+1} has duration {weeks} weeks")

        else:
            print(f"   Duration Warning: Task {tsk} on row {rowIdx+1} is missing a duration")

        if qty is not None and qty != '' and int(qty) > 0 and (asg == '' or asg is None):
            print(f"   Assignment Warning: Task {tsk} on row {rowIdx+1} with {qty} hours")

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


def cost_labor(*, client, div, sheet, rowIdx, cData, laborTable, resourceTable, doCost):

    userMapByName, userMapByEmail = configuration.get_user_map(client=client, div=div)

    if sheet.rows[rowIdx].cells[cData['Task']['position']].value is None or \
       sheet.rows[rowIdx].cells[cData['Task']['position']].value  == '' or \
       sheet.rows[rowIdx].cells[cData['Assigned To']['position']].value is None or \
       sheet.rows[rowIdx].cells[cData['Assigned To']['position']].value  == '':
        return

    # Extract key columns
    if doCost == 'Contingency':
        hours = float(sheet.rows[rowIdx].cells[cData['Contingency Hours']['position']].value)
    else:
        hours = float(sheet.rows[rowIdx].cells[cData['Budgeted Quantity']['position']].value)

    rate = float(sheet.rows[rowIdx].cells[cData['Cost Per Item']['position']].value)
    startStr = sheet.rows[rowIdx].cells[cData['Start']['position']].value
    endStr = sheet.rows[rowIdx].cells[cData['End']['position']].value

    if endStr is None or startStr is None:
        return

    # Add entry to resourceTable is applicable
    resource = sheet.rows[rowIdx].cells[cData['Assigned To']['position']].value

    # Convert email to name
    if resource in userMapByEmail:
        resource = userMapByEmail[resource]

    if resourceTable is not None:
        if resource not in resourceTable:
            resourceTable[resource] = {sheet.name: {}}

        if sheet.name not in resourceTable[resource]:
            resourceTable[resource][sheet.name] = {}

    # Split date fields
    startFields = startStr.split('T')[0].split('-')  # Start date fields
    endFields = endStr.split('T')[0].split('-')      # End date fields

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
    hoursPerDay = 0 if days == 0 else hours / days

    # Iterate through the weekdays
    for n in range(int((endDate - startDate).days)+1):
        dt = startDate + datetime.timedelta(n)

        if dt.weekday() != 5 and dt.weekday() != 6:

            # Generate a key for the year/month
            k = f"{dt.year}_{dt.month:02d}"

            # Add year/month to dictionary if it does not exist
            if k not in laborTable:
                laborTable[k] = {resource: {'rate' : rate, 'hours' : 0.0}}

            # Add rate to sub-dictionrary for month
            elif resource not in laborTable[k]:
                laborTable[k][resource] = {'rate' : rate, 'hours' : 0.0}

            # Increment the number of hours in that year/month, for the given rate
            laborTable[k][resource]['hours'] += hoursPerDay

            if resourceTable is not None:

                # Add year/month to dictionary if it does not exist
                if k not in resourceTable[resource][sheet.name]:
                    resourceTable[resource][sheet.name][k] = 0.0

                # Increment the number of hours in that year/month, for the given rate
                resourceTable[resource][sheet.name][k] += hoursPerDay


def check(*, client, sheet, doFixes, div, cData, doCost, name, doDownload, doTask, resources, resourceTable):
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
    check_row(client=client, sheet=sheet, rowIdx=0, key='top', div=div, cData=cData, doFixes=doFixes, doTask=doTask, resources=None)

    # First walk through the rows and create a parent list
    parents = set()
    for rowIdx in range(len(sheet.rows)):
        parents.add(sheet.rows[rowIdx].parent_id)

    # Process the rest of the rows
    for rowIdx in range(1,len(sheet.rows)):

        # MS Section
        if sheet.rows[rowIdx].cells[0].value == 'M&S':
            check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='ms_top', div=div, cData=cData, doFixes=doFixes, doTask=doTask, resources=None)
            inMS = True
            inLabor = False

        elif sheet.rows[rowIdx].cells[0].value == 'Labor':
            check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='labor_top', div=div, cData=cData, doFixes=doFixes, doTask=doTask, resources=None)
            inMS = False
            inLabor = True

        elif inMS:

            # In parent
            if sheet.rows[rowIdx].id in parents:
                check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='ms_parent', div=div, cData=cData, doFixes=doFixes, doTask=doTask, resources=None)
            else:
                check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='ms_task', div=div, cData=cData, doFixes=doFixes, doTask=doTask, resources=None)
                if doCost != 'None':
                    cost_ms(sheet=sheet, rowIdx=rowIdx, cData=cData, msTable=msTable)

        elif inLabor:

            # In parent
            if sheet.rows[rowIdx].id in parents:
                check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='labor_parent', div=div, cData=cData, doFixes=doFixes, doTask=doTask, resources=None)
            else:
                check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='labor_task', div=div, cData=cData, doFixes=doFixes, doTask=doTask, resources=resources)

                if doCost != 'None' or resourceTable is not None:
                    cost_labor(client=client, div=div, sheet=sheet, rowIdx=rowIdx, cData=cData, laborTable=laborTable, resourceTable=resourceTable, doCost=doCost)

    # Generate excel friend view of monthly spending
    if doCost != 'None':
        write_cost_table(name=name, laborTable=laborTable, msTable=msTable)

    if isinstance(doDownload,str):
        client.Sheets.get_sheet_as_excel(sheet.id, doDownload)

    return True


def write_cost_table(*, name, laborTable, msTable):

    # Generate excel friend view of monthly spending
    months = set([])
    resources = set([])

    # First generate month and rate lists to create excel grid
    for mnth in laborTable:
        months.add(mnth)
        for res in laborTable[mnth]:
            resources.add(res)

    for mnth in msTable:
        months.add(mnth)

    months = sorted(months)
    resources = sorted(resources)

    with open(f'{name} Cost.csv', 'w') as f:

        # Create Header
        f.write('Month\tM&S\t')

        for k in resources:
            f.write(f'{k} Rate\t{k} Hours\t{k} Dollars\t')

        f.write("Total Labor Hours\tTotal Labor Dollars\tTotal Dollars\n")

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

            f.write(f'{m}\t{val:0.2f}\t')
            totDollars += val

            # Put in Labor
            for res in resources:
                if m in laborTable and res in laborTable[m]:
                    rte = float(laborTable[m][res]['rate'])
                    hrs = float(laborTable[m][res]['hours'])
                    val = float(hrs) * float(rte)
                else:
                    rte = 0.0
                    hrs = 0.0
                    val = 0.0

                totLabHours += hrs
                totLabDollars += val
                totDollars += val

                f.write(f'{rte}\t{hrs:0.2f}\t{val:0.2f}\t')

            f.write(f"{totLabHours:0.2f}\t{totLabDollars:0.2f}\t{totDollars:0.2f}\n")


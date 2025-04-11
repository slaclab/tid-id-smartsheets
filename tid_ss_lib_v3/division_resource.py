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
import copy
from . import navigate
from . import tracking_sheet_columns
from . import configuration

# Set formats
#
# https://smartsheet-platform.github.io/api-docs/#formatting
#
# Colors 31 = Dark Blue
#        26 = Dark Gray
#        23 = Blue
#        18 - Gray
#           = White

Months = { 'FY2025': [ '2024_10', '2024_11', '2024_12', '2025_01', '2025_02', '2025_03', '2025_04', '2025_05', '2025_06', '2025_07', '2025_08', '2025_09' ],
           'FY2026': [ '2025_10', '2025_11', '2025_12', '2026_01', '2026_02', '2026_03', '2026_04', '2026_05', '2026_06', '2026_07', '2026_08', '2026_09' ]}

HoursPerMonth = 160
HoursPerYear = 1780

ColData = { 'Resource / Project': { 'position' : 0, 'type': 'TEXT_NUMBER',
                                    'resource' : { 'format': ",,1,,,,,,2,31,,,,,,1,",  'formula' : None },
                                    'project'  : { 'format': None,                     'formula' : None } },

            'Department': { 'position' : 1, 'type': 'TEXT_NUMBER',
                            'resource' : { 'format': ",,1,,,,,,2,31,,,,,,1,",  'formula' : None },
                            'project'  : { 'format': None,                     'formula' : None } }}

idx = 2;

for k in Months:
    ColData[f'{k} Hours'] = { 'position' : idx, 'type': 'TEXT_NUMBER',
                              'resource' : { 'format': ",,1,,,,,,2,31,,,0,,1,1,",  'formula' : '=SUM(CHILDREN())' },
                              'project'  : { 'format': ",,,,,,,,,,,,0,,1,,",       'formula' : None } }
    idx += 1

    ColData[f'{k} Pct'] = { 'position' : idx, 'type': 'TEXT_NUMBER',
                            'resource' : { 'format': ",,1,,,,,,2,31,,,0,1,3,1,",  'formula' : f'=[{k} Hours]@row / {HoursPerYear}' },
                            'project'  : { 'format': ",,,,,,,,,18,,,0,1,3,,",    'formula' : f'=[{k} Hours]@row / {HoursPerYear}' } }
    idx += 1

for k,v in Months.items():
    for m in v:

        ColData[f'{m} Hours'] = { 'position' : idx, 'type': 'TEXT_NUMBER',
                                  'resource' : { 'format': ",,1,,,,,,2,31,,,0,,1,1,",  'formula' : '=SUM(CHILDREN())' },
                                  'project'  : { 'format': ",,,,,,,,,,,,0,,1,,",       'formula' : None } }
        idx += 1

        ColData[f'{m} Pct'] = { 'position' : idx, 'type': 'TEXT_NUMBER',
                                'resource' : { 'format': ",,1,,,,,,2,31,,,0,1,3,1,",  'formula' : f'=[{m} Hours]@row / {HoursPerMonth}' },
                                'project'  : { 'format': ",,,,,,,,,18,,,0,1,3,,",    'formula' : f'=[{m} Hours]@row / {HoursPerMonth}' } }
        idx += 1


def find_columns(*, client, sheet, rData):

    # Look for each expected column
    for k,v in rData.items():
        found = False
        for i in range(len(sheet.columns)):
            if sheet.columns[i].title == k:
                found = True
                if v['position'] != i:
                    print(f"    Resource column location mismatch for {k}. Expected at {v['position']+1}, found at {i+1}.")
                    v['position'] = i

        if not found:
            print(f"    Adding Resource column: {k}.")
            col = smartsheet.models.Column({'title': k,
                                            'type': v['type'],
                                            'index': v['position']})

            client.Sheets.add_columns(sheet.id, [col])
            return False

    return True


def del_rows(*, client, sheet, rData):
    while (len(sheet.rows) > 0 ):

        delCnt = 100 if len(sheet.rows) > 100 else len(sheet.rows)

        print(f"    Deleting {delCnt} rows")

        client.Sheets.delete_rows(sheet.id, [sheet.rows[i].id for i in range(delCnt)])

        sheet = client.Sheets.get_sheet(sheet.id, include='format')


def add_resources(*, client, sheet, rData, resourceData):
    print("    Adding resources")
    addRows = []

    for k,v in resourceData.items():
        new_row = smartsheet.models.Row()

        for ck,cv in rData.items():

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[cv['position']].id

            if ck == 'Resource / Project':
                new_cell.value = v['name']
            elif ck == 'Department':
                new_cell.value = v['dep']
            elif cv['resource']['formula'] is not None:
                new_cell.formula = cv['resource']['formula']
            else:
                new_cell.value = ""

            if cv['resource']['format'] is not None:
                new_cell.format = cv['resource']['format']

            new_cell.strict = False
            new_row.cells.append(new_cell)

        addRows.append(new_row)

    # Process add rows
    if len(addRows) > 0:
        client.Sheets.add_rows(sheet.id, addRows)


def add_projects(*, client, sheet, rData, resourceData):
    print("    Adding projects")

    for r in range(len(sheet.rows)):
        entry = sheet.rows[r].cells[0].value
        addRows = []

        # Each Project
        for proj,projMonths in resourceData[entry]['projects'].items():
            new_row = smartsheet.models.Row()
            new_row.parent_id = sheet.rows[r].id

            # Project Name
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[rData['Resource / Project']['position']].id
            new_cell.value = proj

            if rData['Resource / Project']['project']['format'] is not None:
                new_cell.format = rData['Resource / Project']['project']['format']

            new_cell.strict = False
            new_row.cells.append(new_cell)

            # Department
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[rData['Department']['position']].id
            new_cell.value = resourceData[entry]['dep']

            if rData['Resource / Project']['project']['format'] is not None:
                new_cell.format = rData['Resource / Project']['project']['format']

            new_cell.strict = False
            new_row.cells.append(new_cell)

            for k,v in Months.items():
                totHours = 0

                for month in v:

                    # Add hours cell
                    key = f"{month} Hours"

                    if month in projMonths:
                        value = projMonths[month]
                        totHours += value
                    else:
                        value = 0.0

                    pos = rData[key]['position']

                    new_cell = smartsheet.models.Cell()
                    new_cell.column_id = sheet.columns[rData[key]['position']].id
                    new_cell.value = value

                    if rData[key]['project']['format'] is not None:
                        new_cell.format = rData[key]['project']['format']

                    new_cell.strict = False
                    new_row.cells.append(new_cell)

                    # Add percentage cell
                    key = f"{month} Pct"

                    pos = rData[key]['position']

                    new_cell = smartsheet.models.Cell()
                    new_cell.column_id = sheet.columns[rData[key]['position']].id
                    new_cell.formula = rData[key]['project']['formula']

                    if rData[key]['project']['format'] is not None:
                        new_cell.format = rData[key]['project']['format']

                    new_cell.strict = False
                    new_row.cells.append(new_cell)

                # Total Hours
                key = f"{k} Hours"

                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[rData[key]['position']].id
                new_cell.value = totHours

                if rData[key]['project']['format'] is not None:
                    new_cell.format = rData[key]['project']['format']

                new_cell.strict = False
                new_row.cells.append(new_cell)

                # Total Pct
                key = f"{k} Pct"

                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[rData[key]['position']].id
                new_cell.formula = rData[key]['project']['formula']

                if rData[key]['project']['format'] is not None:
                    new_cell.format = rData[key]['project']['format']

                new_cell.strict = False
                new_row.cells.append(new_cell)

            addRows.append(new_row)

        # Process add rows
        if len(addRows) > 0:
            client.Sheets.add_rows(sheet.id, addRows)


def sort_resources(*, client, div, resourceTable):
    print("    Sorting and lookup up users")

    userMapByEmail = configuration.get_user_map_by_email(client=client, div=div)
    resourceData = {}

    # Reformat resource table, adding department information
    for entry, projects in resourceTable.items():
        name = entry
        dep  = "Unknown"

        # Most SM entries are emails
        if entry in userMapByEmail:
            name = userMapByEmail[entry]['name']

            if userMapByEmail[entry]['dep'] is not None:
                dep  = userMapByEmail[entry]['dep']

        if name in resourceData:
            print(f"    Found duplicate name {name}")
            name += " Duplicate"

        resourceData[name] = {'dep' : dep,
                              'name' : name,
                              'entry' : entry,
                              'projects' : projects}

    # Sort the dictionary by dept then name
    sortedData = dict(sorted(resourceData.items(), key=lambda item: item[1]['dep'] + item[1]['name']))

    return sortedData


def update(*, client, div, resourceTable):
    print("Updating division resource table")
    rData = copy.deepcopy(ColData)
    sheet = client.Sheets.get_sheet(int(div.division_resources), include='format')

    resourceData = sort_resources(client=client, div=div, resourceTable=resourceTable)

    # First check structure
    while True:
        ret = find_columns(client=client, sheet=sheet, rData=rData )
        if ret is True:
            break
        else:
            sheet = client.Sheets.get_sheet(sheet.id, include='format')

    # Next delete rows
    del_rows(client=client, sheet=sheet, rData=rData)
    sheet = client.Sheets.get_sheet(sheet.id, include='format')

    # Next add resources
    add_resources(client=client, sheet=sheet, rData=rData, resourceData=resourceData)
    sheet = client.Sheets.get_sheet(sheet.id, include='format')

    # Next add projects
    add_projects(client=client, sheet=sheet, rData=rData, resourceData=resourceData)


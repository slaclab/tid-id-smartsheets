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
from . import configuration

import datetime
import copy

ColDataActuals = { 'Lookup PA'          : {'position' : 0},
                   'Monthly Actuals'    : {'position' : 2},
                   'Total Actuals'      : {'position' : 3},
                   'Actuals Start Date' : {'position' : 5},
                   'Project Folder ID'  : {'position' : 7}}

ColDataWbs = { 'PA Proj_Act'        : {'position' : 2},  # XXXXX-YYYY
               'Period-Month'       : {'position' : 4},  # 01-OCT
               'Employee Name'      : {'position' : 6},  # text name
               'Values Hrs ACTUALS' : {'position' : 7},  # hours
               'Cost ACTUALS'       : {'position' : 9}}  # $xxxx.xx

def find_columns(*, client, sheet, cData):

    # Look for each expected column
    for k,v in cData.items():
        found = False
        for i in range(len(sheet.columns)):
            if sheet.columns[i].title == k:
                found = True
                if v['position'] != i:
                    #print(f"   Division actuals or wbs column location mismatch for {k}. Expected at {v['position']+1}, found at {i+1}.")
                    v['position'] = i

        if not found:
            print(f"   Division actuals or wbs column not found: {k}.")
            return False

    return True


# Returns two dictionaries
#
# The first is a dictionary with the PA number as the key, each entry containing the following fields:
#    'Lookup PA'          = PA Number xxxxx-yyyy
#    'Project Folder ID'  = api id for sheet
#    'Monthly Actuals'    = $xxxx.xx
#    'Total Actuals'      = $xxxx.xx
#    'Actuals Start Date' = YYYY-MM-DD
#
# The second is a dictionary with the project id as the key, each containing the following fields:
#   'pas'    : {} = Dictionary of PA dictionaries associated with this project (see above structure), PA number is key
#   'months' : [] = list of months which will be included in the actuals YYYY_MM format
#   'person' : {} = Dictionary of people who have charged to this project, key is the name format as
#                   included in the WBS file. Each entry is a dictionary with the following format.
#                   (inititally empty as no persons are added initially)
#            'total_cost'    = Total amount charged to the project
#            'total_hours'   = Total amount charged to the project
#            'monthly_hours' = Hours by month, dictionary with YY_MM as key
#            'monthly_cost'  = Cost by month, dictionary with YY_MM as key
#
def get_div_actuals_data (*, client, div):

    cData = copy.deepcopy(ColDataActuals)

    sheet = client.Sheets.get_sheet(int(div.division_actuals))

    find_columns(client=client, sheet=sheet, cData=cData)

    retPaData = {}
    retProjData = {}

    lastYear = int(div.earned_value_date.split('-')[0])
    lastMonth = int(div.earned_value_date.split('-')[1])

    if lastMonth >= 10:
        lastYear += 1

    # Process the rows
    for rowIdx in range(0,len(sheet.rows)):
        entry = {}

        for k,v in cData.items():
            entry[k] = sheet.rows[rowIdx].cells[v['position']].value;

        retPaData[sheet.rows[rowIdx].cells[cData['Lookup PA']['position']].value] = entry;

        fid = entry['Project Folder ID']

        if fid not in retProjData:
            retProjData[fid] = { 'pas' : {}, 'months' : [], 'person' : {} }

        retProjData[fid]['pas'][entry['Lookup PA']] = entry

    # Generate date list per project
    for k,v in retProjData.items():
        firstPa = next(iter(v['pas']))

        # Set start date for actuals, use first dictionary entry
        year  = int(v['pas'][firstPa]['Actuals Start Date'].split('-')[0])
        month = int(v['pas'][firstPa]['Actuals Start Date'].split('-')[1])
        v['months'].append(f"{year:04}_{month:02}")

        while True:
            if  year == lastYear and month == lastMonth:
                break;
            elif month == 9:
                month = 10
                year += 1
            elif month == 12:
                month = 1
            else:
                month += 1

            v['months'].append(f"{year:04}_{month:02}")

    return retPaData, retProjData


def parse_wbs_actuals_sheet(*, client, div, sheetId, year, paData, projData):
    cData = copy.deepcopy(ColDataWbs)

    sheet = client.Sheets.get_sheet(sheetId)

    find_columns(client=client, sheet=sheet, cData=cData)

    userMap = configuration.get_user_map(client=client, div=div)
    missMap = set()

    # Process the rows
    for rowIdx in range(0,len(sheet.rows)):
        entry = {}

        for k,v in cData.items():
            entry[k] = sheet.rows[rowIdx].cells[v['position']].value;

        pa = entry['PA Proj_Act']

        if pa not in paData:
            #print(f"Could not find PA number {pa} found in {year} in in actuals lookup sheet")
            pass

        elif entry['Period-Month'] is not None and entry['Employee Name'] is not None and \
             entry['Values Hrs ACTUALS'] is not None and entry['Cost ACTUALS'] is not None:

            pid = paData[pa]['Project Folder ID']

            pData = projData[pid]

            month = int(entry['Period-Month'].split('-')[0])

            if month < 4:
                month += 9
            else:
                month -= 3

            name  = entry['Employee Name']
            hours = float(entry['Values Hrs ACTUALS'])
            cost  = float(entry['Cost ACTUALS'])
            dStr = f"{year:04}_{month:02}"

            if name in userMap:
                email = userMap[name]
            else:
                print(f"User name {name} not found in email map")
                missMap.add(name)
                email = name

            if email not in pData['person']:
                pData['person'][email] = {'total_cost' : cost,
                                          'total_hours'  : hours,
                                          'monthly_hours' : { dStr : hours},
                                          'monthly_cost'  : { dStr : cost}}
            else:
                pData['person'][email]['total_cost'] += cost
                pData['person'][email]['total_hours'] += hours
                pData['person'][email]['monthly_hours'][dStr] = hours
                pData['person'][email]['monthly_cost'][dStr] = cost

    configuration.add_missing_user_map(client=client, div=div, miss=missMap)


def get_wbs_actuals(*, client, div):

    if div.wbs_exports is None:
        return None

    retPaData, retProjData = get_div_actuals_data (client=client, div=div)

    for k,v in div.wbs_exports.items():
        parse_wbs_actuals_sheet(client=client, div=div, sheetId=int(v), year=k, paData=retPaData, projData=retProjData)

    return retProjData


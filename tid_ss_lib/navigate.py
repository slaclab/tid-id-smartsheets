#-----------------------------------------------------------------------------
# Title      : Top Level Navigation Budget Sheet
#-----------------------------------------------------------------------------
# This file is part of the TID ID Smartsheets software platform. It is subject to
# the license terms in the LICENSE.txt file found in the top-level directory
# of this distribution and at:
#    https://confluence.slac.stanford.edu/display/ppareg/LICENSE.html.
# No part of the TID ID Smartsheets software platform, including this file, may be
# copied, modified, propagated, or distributed except according to the terms
# contained in the LICENSE.txt file.
#-----------------------------------------------------------------------------

import smartsheet  # pip3 install smartsheet-python-sdk
from . import budget_sheet
from . import schedule_sheet
from . import tracking_sheet

import numpy as np
import datetime

TID_WORKSPACE = 4728845933799300
OVERHEAD_NOTE = '12.25% Overhead'

TID_ID_ACTIVE_FOLDER     = 1039693589571460
TID_ID_LIST_SHEET        = 2931334483076996
TID_ID_FOLDER_PREFIX     = 'TID/ID'
TID_ID_ACTUALS_SHEET     = 7403570111768452
TID_ID_ACTUALS_START_ROW = 5693264191481732
TID_ID_ACTUALS_END_ROW   = 1256792717715332
TID_ID_RESOURCE_FOLDER   = 6665944920549252
TID_ID_TEMPLATE_FOLDER   = 5079595864090500
TID_ID_TEMPLATE_BUDGET   = 1091860589504388
TID_ID_TEMPLATE_SCORING  = 3343660403189636
TID_ID_TEMPLATE_RISK     = 7847260030560132
TID_ID_TEMPLATE_SCHEDULE = 8410209983981444
TID_ID_TEMPLATE_TRACKING = 5595460216874884

5079595864090500

TID_ID_RATE_NOTE = 'TID-ID Eng Rate FY23 $291; Tech Rate FY23: $167'

TID_CDS_ACTIVE_FOLDER     = 8506630263334788
TID_CDS_LIST_SHEET        = 4947269792360324
TID_CDS_FOLDER_PREFIX     = 'TID/CDS'
TID_CDS_ACTUALS_SHEET     = 2695469978675076
TID_CDS_ACTUALS_START_ROW = 3233931241645956
TID_CDS_ACTUALS_END_ROW   = 8729290357270404
TID_CDS_RESOURCE_FOLDER   = 7728176030869380
TID_CDS_TEMPLATE_FOLDER   = 3048654543054724
TID_CDS_TEMPLATE_BUDGET   = 1489735219734404
TID_CDS_TEMPLATE_SCORING  = 3741535033419652
TID_CDS_TEMPLATE_RISK     = 8245134660790148
TID_CDS_TEMPLATE_SCHEDULE = 7119234753947524
TID_CDS_TEMPLATE_TRACKING = 5993334847104900

TID_CDS_RATE_NOTE = 'TID-CDS Eng Rate FY23 $291; Tech Rate FY23: $167'

OVERHEAD_NOTE = '12.25% Overhead'


def gen_standard_sheets(*, div):

    if div == 'id':
        return {'Budget': TID_ID_TEMPLATE_BUDGET,
                'PM Scoring': TID_ID_TEMPLATE_SCORING,
                'Risk Registry': TID_ID_TEMPLATE_RISK,
                'Schedule': TID_ID_TEMPLATE_SCHEDULE,
                'Tracking': TID_ID_TEMPLATE_TRACKING}
    elif div == 'cds':
        return {'Budget': TID_CDS_TEMPLATE_BUDGET,
                'PM Scoring': TID_CDS_TEMPLATE_SCORING,
                'Risk Registry': TID_CDS_TEMPLATE_RISK,
                'Schedule': TID_CDS_TEMPLATE_SCHEDULE,
                'Tracking': TID_CDS_TEMPLATE_TRACKING}


def get_folder_data(*, client, div, folderId, path=None):
    folder = client.Folders.get_folder(folderId)

    ret = {'folder': folder}

    ret['path'] = path
    ret['tracked'] = False
    ret['name'] = folder.name
    ret['url'] = folder.permalink
    ret['sheets'] = {k: None for k in gen_standard_sheets(div=div)}

    for s in folder.sheets:
        for k in gen_standard_sheets(div=div):
            if k == s.name[-len(k):]:
                ret['sheets'][k] = s

    return ret


def check_project(*, client, div, folderId, doFixes, doCost=False, path=None):
    fdata = get_folder_data(client=client, div=div, folderId=folderId)
    laborRows = None

    if path is not None:
        print(f"Processing project {path} : {folderId}")
    else:
        print(f"Processing project {fdata['folder'].name} : {folderId}")

    ##########################################################
    # First Make sure folder has all of the neccessary files
    ##########################################################
    for k, v in fdata['sheets'].items():

        # Copy file if it is missing
        if v is None:
            print(f"   Project is missing '{k}' file.")

            if doFixes:
                print(f"   Coping '{k}' file to project.")
                client.Sheets.copy_sheet(StandardProjectFiles[k], # Source sheet
                                         smartsheet.models.ContainerDestination({'destination_type': 'folder',
                                                                                'destination_id': fdata['folder'].id,
                                                                                'new_name': fdata['folder'].name + ' ' + k}))

        # Check for valid naming, rename if need be
        elif 'Template Set ' not in fdata['folder'].name and not v.name.startswith(fdata['folder'].name):
            print(f"   Bad sheet name {v.name}.")

            if doFixes:
                print(f"   Renaming {v.name}.")
                client.Sheets.update_sheet(v.id, smartsheet.models.Sheet({'name': fdata['folder'].name + ' ' + k}))

    # Refresh folder data, needed if new files were copied over
    fdata = get_folder_data(client=client, div=div, folderId=folderId)

    if fdata['sheets']['Budget'] is None or fdata['sheets']['Schedule'] is None:
        print("   Skipping remaining processing")
        return

    # Re-read sheet data
    fdata['sheets']['Budget'] = client.Sheets.get_sheet(fdata['sheets']['Budget'].id, include='format')
    fdata['sheets']['Schedule'] = client.Sheets.get_sheet(fdata['sheets']['Schedule'].id, include='format')
    fdata['sheets']['Tracking'] = client.Sheets.get_sheet(fdata['sheets']['Tracking'].id, include='format')

    # Double check budget for new fix
    if doFixes and not budget_sheet.check_structure(sheet=fdata['sheets']['Budget']):
        print("   Attempting to update budget sheet")
        budget_sheet.fix_structure(client=client, sheet=fdata['sheets']['Budget'])
        fdata['sheets']['Budget'] = client.Sheets.get_sheet(fdata['sheets']['Budget'].id, include='format')

    # Double check schedule for new fix
    if doFixes and not schedule_sheet.check_structure(sheet=fdata['sheets']['Schedule']):
        print("   Attempting to update schedule sheet")
        schedule_sheet.fix_structure(client=client, sheet=fdata['sheets']['Schedule'])
        fdata['sheets']['Schedule'] = client.Sheets.get_sheet(fdata['sheets']['Schedule'].id, include='format')

    # Double check tracking for new fix
    if doFixes and not tracking_sheet.check_structure(sheet=fdata['sheets']['Tracking']):
        print("   Attempting to update tracking sheet")
        tracking_sheet.fix_structure(client=client, div=div, sheet=fdata['sheets']['Tracking'])
        fdata['sheets']['Tracking'] = client.Sheets.get_sheet(fdata['sheets']['Tracking'].id, include='format')

    if budget_sheet.check_structure(sheet=fdata['sheets']['Budget']) and schedule_sheet.check_structure(sheet=fdata['sheets']['Schedule']) and tracking_sheet.check_structure(sheet=fdata['sheets']['Tracking']):

        # Fix internal budget file references
        laborRows = budget_sheet.check(client=client, div=div, sheet=fdata['sheets']['Budget'], doFixes=doFixes )

        # Check schedule file
        schedule_sheet.check(client=client, sheet=fdata['sheets']['Schedule'], laborRows=laborRows, laborSheet=fdata['sheets']['Budget'], doFixes=doFixes )

        # Final fix of links in budget file
        budget_sheet.check_task_links(client=client, sheet=fdata['sheets']['Budget'], laborRows=laborRows, scheduleSheet=fdata['sheets']['Schedule'], doFixes=doFixes)

        # Fix tracking file
        tracking_sheet.check(client=client, sheet=fdata['sheets']['Tracking'], budgetSheet=fdata['sheets']['Budget'], doFixes=doFixes)

    else:
        print("   Skipping remaining processing")

    if laborRows is not None and doCost is True:
        compute_monthly_cost(name = fdata['folder'].name, data=laborRows)


def compute_monthly_cost(*, name, data, fy=2023):

    tot = 0.0
    months = {f'{fy-1}_10': 0.0, f'{fy-1}_11': 0.0, f'{fy-1}_12': 0.0,
              f'{fy}_1': 0.0, f'{fy}_2': 0.0, f'{fy}_3': 0.0,
              f'{fy}_4': 0.0, f'{fy}_5': 0.0, f'{fy}_6': 0.0,
              f'{fy}_7': 0.0, f'{fy}_8': 0.0, f'{fy}_9': 0.0}

    for r in data:
        sd = r['link'].cells[5].value.split('T')[0].split('-')  # Start date fields
        ed = r['link'].cells[6].value.split('T')[0].split('-')  # End date fields

        sdd = datetime.date(int(sd[0]),int(sd[1]),int(sd[2]))
        edd = datetime.date(int(ed[0]),int(ed[1]),int(ed[2]))

        if r['parent'] is False:
            days = 0

            # Count the number of weekdays in the time period
            for n in range(int((edd - sdd).days)):
                dt = sdd + datetime.timedelta(n)

                if dt.weekday() != 5 and dt.weekday() != 6:
                    days += 1

            # (Hours * rate) / days to get cost per day
            cpd = (float(r['data'].cells[3].value) * float(r['data'].cells[2].value)) / days

            # Iterate through the days
            for n in range(int((edd - sdd).days)):
                dt = sdd + datetime.timedelta(n)

                if dt.weekday() != 5 and dt.weekday() != 6:
                    k = f"{dt.year}_{dt.month}"

                    if k in months:
                        months[k] += cpd
                        tot += cpd

    with open(f'{name}_cost.txt', 'w') as f:
        for k,v in months.items():
            f.write(f"{k},{v}\n")

def get_active_list(*, client, div, path=None, folderId=None):

    if path is None and folderId is None:
        if div == 'id':
            path     = TID_ID_FOLDER_PREFIX
            folderId = TID_ID_ACTIVE_FOLDER
        elif div == 'cds':
            path     = TID_CDS_FOLDER_PREFIX
            folderId = TID_CDS_ACTIVE_FOLDER

    folder = client.Folders.get_folder(folderId)
    ret = {}

    path = path + '/' + folder.name

    # No sub folders, this might be a project
    if len(folder.folders) == 0:

        # Skip projects with no sheets
        if len(folder.sheets) != 0:
            ret[folderId] = get_folder_data(client=client, div=div, folderId=folder.id, path=path)

    else:
        for sub in folder.folders:
            ret.update(get_active_list(client=client, div=div, path=path, folderId=sub.id))

    return ret

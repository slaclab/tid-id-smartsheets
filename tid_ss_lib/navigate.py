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

TID_WORKSPACE          = 4728845933799300
TID_ID_ACTIVE_FOLDER   = 1039693589571460
TID_ID_TEMPLATE_FOLDER = 4013014891423620
TID_ID_LIST_SHEET      = 2931334483076996
TID_ID_FOLDER_PREFIX   = 'TID/ID'
TID_ACTUALS_SHEET      = 7403570111768452
TID_ACTUALS_START_ROW  = 5693264191481732
TID_ACTUALS_END_ROW    = 1256792717715332
TID_RESOURCE_FOLDER    = 6665944920549252

OVERHEAD_NOTE = '12.25% Overhead'
LABOR_RATE_NOTE = 'Labor Rate FY22 (Oct - Feb): $273.25; (Mar - Sep) $281.45; Slac Tech Rate FY22: $162.65'

# Standard Project Files & Template IDs
StandardProjectFiles = {'Budget': 6232441649162116,
                        'PM Scoring': 888744673863556,
                        'Risk Registry': 607544575059844,
                        'Schedule': 1728842021791620,
                        'Tracking': 4586297051375492}


def get_folder_data(*, client, folderId, path=None):
    folder = client.Folders.get_folder(folderId)

    ret = {'folder': folder}

    ret['path'] = path
    ret['tracked'] = False
    ret['name'] = folder.name
    ret['url'] = folder.permalink

    for k in StandardProjectFiles:
        ret[k] = None

    for s in folder.sheets:
        for k in StandardProjectFiles:
            if k == s.name[-len(k):]:
                ret[k] = s

    return ret


def check_project(*, client, folderId, doFixes, path=None):
    fdata = get_folder_data(client=client, folderId=folderId)

    if path is not None:
        print(f"Processing project {path} : {folderId}")
    else:
        print(f"Processing project {fdata['folder'].name} : {folderId}")

    ##########################################################
    # First Make sure folder has all of the neccessary files
    ##########################################################
    for k, v in fdata.items():

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
    fdata = get_folder_data(client=client, folderId=folderId)

    if fdata['Budget'] is None or fdata['Schedule'] is None:
        print("   Skipping remaining processing")
        return

    # Re-read sheet data
    fdata['Budget'] = client.Sheets.get_sheet(fdata['Budget'].id, include='format')
    fdata['Schedule'] = client.Sheets.get_sheet(fdata['Schedule'].id, include='format')
    fdata['Tracking'] = client.Sheets.get_sheet(fdata['Tracking'].id, include='format')

    # Double check schedule for new fix
    if doFixes and not schedule_sheet.check_structure(sheet=fdata['Schedule']):
        print("   Attempting to update schedule sheet")
        schedule_sheet.fix_structure(client=client, sheet=fdata['Schedule'])
        fdata['Schedule'] = client.Sheets.get_sheet(fdata['Schedule'].id, include='format')

    # Double check tracking for new fix
    if doFixes and not tracking_sheet.check_structure(sheet=fdata['Tracking']):
        print("   Attempting to update tracking sheet")
        tracking_sheet.fix_structure(client=client, sheet=fdata['Tracking'])
        fdata['Tracking'] = client.Sheets.get_sheet(fdata['Tracking'].id, include='format')

    if budget_sheet.check_structure(sheet=fdata['Budget']) and schedule_sheet.check_structure(sheet=fdata['Schedule']) and tracking_sheet.check_structure(sheet=fdata['Tracking']):

        # Fix internal budget file references
        laborRows = budget_sheet.check(client=client, sheet=fdata['Budget'], doFixes=doFixes )

        # Check schedule file
        schedule_sheet.check(client=client, sheet=fdata['Schedule'], laborRows=laborRows, laborSheet=fdata['Budget'], doFixes=doFixes )

        # Final fix of links in budget file
        budget_sheet.check_task_links(client=client, sheet=fdata['Budget'], laborRows=laborRows, scheduleSheet=fdata['Schedule'], doFixes=doFixes)

        # Fix tracking file
        tracking_sheet.check(client=client, sheet=fdata['Tracking'], budgetSheet=fdata['Budget'], doFixes=doFixes)

    else:
        print("   Skipping remaining processing")


def get_active_list(*, client, path = TID_ID_FOLDER_PREFIX, folderId=TID_ID_ACTIVE_FOLDER):
    folder = client.Folders.get_folder(folderId)
    ret = {}

    path = path + '/' + folder.name

    # No sub folders, this might be a project
    if len(folder.folders) == 0:

        # Skip projects with no sheets
        if len(folder.sheets) != 0:
            ret[folderId] = get_folder_data(client=client, folderId=folder.id, path=path)

    else:
        for sub in folder.folders:
            ret.update(get_active_list(client=client, path=path, folderId=sub.id))

    return ret

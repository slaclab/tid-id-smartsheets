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

import smartsheet
from . import budget_sheet
from . import schedule_sheet

TID_WORKSPACE          = 4728845933799300
TID_ID_ACTIVE_FOLDER   = 1039693589571460
TID_ID_TEMPLATE_FOLDER = 4013014891423620
TID_ID_LIST_SHEET      = 2931334483076996
TID_ID_FOLDER_PREFIX   = 'TID/ID'

OVERHEAD_NOTE = '12.25% Overhead'
LABOR_RATE_NOTE = 'SLAC Labor Rate FY21 (loaded): $274.63; Slac Tech Rate FY21 (loaded): $163.47'

TEMPLATE_PREFIX = '[Project] '

def build_template(*, client):
    temp = {}

    folder = client.Folders.get_folder(TID_ID_TEMPLATE_FOLDER)

    for f in folder.sheets:
        name = f.name[len(TEMPLATE_PREFIX):]
        temp[name] = f.id

    return temp


def check_project(*, client, folderId, doFixes, path=None):
    folder = client.Folders.get_folder(folderId)

    if path is not None:
        print(f"Processing project {path} : {folderId}")
    else:
        print(f"Processing project {folder.name} : {folderId}")

    ##########################################################
    # First Make sure folder has all of the neccessary files
    ##########################################################
    tempList = build_template(client=client)
    foundList = {k: None for k in tempList}

    for s in folder.sheets:
        for k in foundList:
            if k in s.name:
                foundList[k] = s

    for k,v in foundList.items():

        # Copy file if it is missing
        if v is None:
            print(f"   Project is missing '{k}' file.")

            if doFixes:
                print(f"   Coping '{k}' file to project.")
                client.Sheets.copy_sheet(tempList[k], # Source sheet
                                         smartsheet.models.ContainerDestination({'destination_type': 'folder',
                                                                                'destination_id': folder.id,
                                                                                'new_name': folder.name + ' ' + k}))

        # Check for valid naming, rename if need be
        elif not v.name.startswith(folder.name):
            print(f"   Bad sheet name {v.name}.")

            if doFixes:
                print(f"   Renaming {v.name}.")
                client.Sheets.update_sheet(v.id, smartsheet.models.Sheet({'name': folder.name + ' ' + k}))

    if foundList['Budget'] is None or foundList['Schedule'] is None:
        print("   Skipping remaining processing")
    else:

        # First process budget sheet:
        budget   = client.Sheets.get_sheet(foundList['Budget'].id)
        schedule = client.Sheets.get_sheet(foundList['Schedule'].id)

        if budget_sheet.check_structure(sheet=budget) and schedule_sheet.check_structure(sheet=schedule):

            # Fix internal budget file references
            laborRows = budget_sheet.check(client=client, sheet=budget, doFixes=doFixes)

            # Check schedule file
            schedule_sheet.check(client=client, sheet=schedule, laborRows=laborRows, laborSheet=budget, doFixes=doFixes)

            # Final fix of links in budget file
            budget_sheet.check_task_links(client=client, sheet=budget, laborRows=laborRows, scheduleSheet=schedule, doFixes=doFixes)
        else:
            print("   Skipping remaining processing")


def check_folders(*, client, path = TID_ID_FOLDER_PREFIX, folderId=TID_ID_ACTIVE_FOLDER, doFixes):
    folder = client.Folders.get_folder(folderId)
    ret = {}

    path = path + '/' + folder.name

    # No sub folders, this might be a project
    if len(folder.folders) == 0:

        # Skip projects with no sheets
        if len(folder.sheets) != 0:
            ret[path] = folderId
            check_project(client=client, folderId=folderId, doFixes=doFixes, path=path)

    else:
        for sub in folder.folders:
            ret.update(check_folders(client=client, path=path, folderId=sub.id, doFixes=doFixes,))

    return ret

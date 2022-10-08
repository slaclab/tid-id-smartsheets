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
from . import project_sheet
from . import project_sheet_columns
from . import project_list
from . import tracking_sheet

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

TID_ID_RATE_NOTE = 'TID-ID Eng Rate FY23 $291; Tech Rate FY23: $167'

TID_CDS_ACTIVE_FOLDER     = 8506630263334788
TID_CDS_LIST_SHEET        = 4947269792360324
TID_CDS_FOLDER_PREFIX     = 'TID/CDS'
TID_CDS_ACTUALS_SHEET     = 2695469978675076
TID_CDS_ACTUALS_START_ROW = 3233931241645956
TID_CDS_ACTUALS_END_ROW   = 8729290357270404
TID_CDS_RESOURCE_FOLDER   = 7728176030869380
TID_CDS_TEMPLATE_FOLDER   = 2556056422377348

TID_CDS_RATE_NOTE = 'TID-CDS Eng Rate FY23 $291; Tech Rate FY23: $167'

OVERHEAD_NOTE = '12.25% Overhead'


def get_folder_data(*, client, div, folderId, path=None):
    folder = client.Folders.get_folder(folderId)

    StandardSheets = ['Project', 'PM Scoring', 'Risk Registry', 'Tracking']

    ret = {'folder': folder}

    ret['path'] = path
    ret['tracked'] = False
    ret['name'] = folder.name
    ret['url'] = folder.permalink
    ret['sheets'] = {k: None for k in StandardSheets}

    for s in folder.sheets:
        for k in StandardSheets:
            if k == s.name[-len(k):]:
                ret['sheets'][k] = s

    return ret


def check_project(*, client, div, folderId, doFixes, doCost=False, path=None):
    fdata = get_folder_data(client=client, div=div, folderId=folderId)

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
            return

        # Check for valid naming, rename if need be
        elif 'Template Set ' not in fdata['folder'].name and not v.name.startswith(fdata['folder'].name):
            print(f"   Bad sheet name {v.name}.")

            if doFixes:
                print(f"   Renaming {v.name}.")
                client.Sheets.update_sheet(v.id, smartsheet.models.Sheet({'name': fdata['folder'].name + ' ' + k}))
            else:
                return

    # Refresh folder data, needed if new files were copied over
    fdata = get_folder_data(client=client, div=div, folderId=folderId)

    # Re-read sheet data
    fdata['sheets']['Project'] = client.Sheets.get_sheet(fdata['sheets']['Project'].id, include='format')
    fdata['sheets']['Tracking'] = client.Sheets.get_sheet(fdata['sheets']['Tracking'].id, include='format')

    # Used to store column addresses
    cData = project_sheet_columns.ColData

    # Check project file
    ret = project_sheet.check(client=client, div=div, sheet=fdata['sheets']['Project'], doFixes=doFixes, cData=cData, doCost=doCost, name=fdata['folder'].name )

    # Fix tracking file
    if ret:
        tracking_sheet.check(client=client, sheet=fdata['sheets']['Tracking'], projectSheet=fdata['sheets']['Project'], div=div, doFixes=doFixes, cData=cData)

    else:
        print("   Skipping remaining processing")


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


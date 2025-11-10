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
from . import tracking_sheet_columns
from . import actuals_sheet

import datetime
import copy
import time


def get_folder_data(*, client, div, folderId, path=None):

    folder_meta = client.Folders.get_folder_metadata(folderId)

    folder_sheets = client.Folders.get_folder_children(folderId,children_resource_types=['sheets'])
    folder_reports = client.Folders.get_folder_children(folderId,children_resource_types=['reports'])
    folder_sights = client.Folders.get_folder_children(folderId,children_resource_types=['sights'])

    StandardSheets  = ['Project', 'Tracking', 'Actuals', 'PM Scoring', 'Risk Registry']
    StandardSights  = ['Dashboard']
    StandardReports = ['Report']

    ret = {'folder_meta': folder_meta}

    ret['path'] = path
    ret['tracked'] = False
    ret['name'] = folder_meta.name
    ret['url'] = folder_meta.permalink
    ret['sheets'] = {k: None for k in StandardSheets}
    ret['sights'] = {k: None for k in StandardSights}
    ret['reports'] = {k: None for k in StandardReports}

    for s in folder_sheets.data:
        for k in StandardSheets:
            if k == s.name[-len(k):]:
                ret['sheets'][k] = s

    for s in folder_reports.data:
        for k in StandardReports:
            if k == s.name[-len(k):]:
                ret['reports'][k] = s

    for s in folder_sights.data:
        for k in StandardSights:
            if k == s.name[-len(k):]:
                ret['sights'][k] = s

    return ret


def check_project(*, client, div, folderId, doFixes, doCost=None, doDownload=False, path=None, doTask=False, resourceTable=None):
    if folderId == 0:
        return

    fdata = get_folder_data(client=client, div=div, folderId=folderId)

    if path is not None:
        print(f"Processing project {path} : {folderId}")
    else:
        print(f"Processing project {fdata['folder_meta'].name} : {folderId}")

    ##########################################################
    # First Make sure folder has all of the neccessary files
    ##########################################################
    for k, v in fdata['sheets'].items():

        if v is None:
            print(f"   Project is missing '{k}' file.")
            return

        # Check for valid naming, rename if need be
        elif 'Template Set ' not in fdata['folder_meta'].name and not v.name.startswith(fdata['folder_meta'].name):
            print(f"   Bad sheet name {v.name}.")

            if doFixes:
                print(f"   Renaming {v.name}.")
                client.Sheets.update_sheet(v.id, smartsheet.models.Sheet({'name': fdata['folder_meta'].name + ' ' + k}))
            else:
                return

    for k, v in fdata['reports'].items():

        if v is None:
            print(f"   Project is missing '{k}' file.")
            return

        # Check for valid naming, rename if need be
        elif 'Template Set ' not in fdata['folder_meta'].name and not v.name.startswith(fdata['folder_meta'].name):
            print(f"   Bad report name {v.name}. Please rename manually.")
            return

    for k, v in fdata['sights'].items():

        if v is None:
            print(f"   Project is missing '{k}' file.")
            return

        # Check for valid naming, rename if need be
        elif 'Template Set ' not in fdata['folder_meta'].name and not v.name.startswith(fdata['folder_meta'].name):
            print(f"   Bad sight name {v.name}. Please rename manually.")
            return

    # Refresh folder data, needed if new files were copied over
    fdata = get_folder_data(client=client, div=div, folderId=folderId)

    # Re-read sheet data
    fdata['sheets']['Project'] = client.Sheets.get_sheet(fdata['sheets']['Project'].id, include='format')
    fdata['sheets']['Tracking'] = client.Sheets.get_sheet(fdata['sheets']['Tracking'].id, include='format')
    fdata['sheets']['Actuals'] = client.Sheets.get_sheet(fdata['sheets']['Actuals'].id, include='format')

    # Used to store column addresses
    cData = copy.deepcopy(project_sheet_columns.ColData)
    tData = copy.deepcopy(tracking_sheet_columns.ColData)

    # Check project file
    resources = set()
    ret = project_sheet.check(client=client, div=div, sheet=fdata['sheets']['Project'], doFixes=doFixes, cData=cData, doCost=doCost, name=fdata['folder_meta'].name, doDownload=doDownload, doTask=doTask, resources=resources, resourceTable=resourceTable)

    # Fix tracking file
    if ret:

        # Update actuals file
        actuals_sheet.check(client=client, sheet=fdata['sheets']['Actuals'], doFixes=doFixes, resources=resources)

        # Reload project file
        fdata['sheets']['Project'] = client.Sheets.get_sheet(fdata['sheets']['Project'].id, include='format')

        tracking_sheet.check(client=client, sheet=fdata['sheets']['Tracking'], projectSheet=fdata['sheets']['Project'], actualsSheet=fdata['sheets']['Actuals'], div=div, doFixes=doFixes, cData=cData, tData=tData, doDownload=doDownload)

    else:
        print("   Skipping remaining processing")


def get_active_list(*, client, div):
    path = div.folder_prefix
    folderId = int(div.active_folder)
    ret = {}

    top_folders = client.Folders.get_folder_children(folderId,children_resource_types=['folders'])

    for sub in top_folders.data:
        sub_meta = client.Folders.get_folder_metadata(sub.id)
        sub_path = path + f"/{sub_meta.name}"
        print(f"Processing sub directory {sub_path}")

        sub_folders = client.Folders.get_folder_children(sub.id,children_resource_types=['folders'])

        for proj in sub_folders.data:
            proj_meta = client.Folders.get_folder_metadata(proj.id)
            proj_path = sub_path + f"/{proj_meta.name}"

            print(f"Processing project directory {proj_path}")

            try:
                ret[proj_meta.id] = get_folder_data(client=client, div=div, folderId=proj_meta.id, path=proj_path)
            except Exception as msg:
                print("!!!!!!!!!!!!!!!!!!!!! Error Getting Folder Data !!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                print(f"    Id = {folder_meta.id}")
                print(f"    {msg}")
                print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")

    return ret


def update_project_actuals(*, client, div, folderId, wbsData):
    folderId=int(folderId)

    fdata = get_folder_data(client=client, div=div, folderId=folderId)

    if fdata['sheets']['Actuals'] is None:
        print(f"Skipping Project {fdata['folder_meta'].name}  which is missing its actuals sheet")
        return

    # Re-read sheet data
    fdata['sheets']['Actuals'] = client.Sheets.get_sheet(fdata['sheets']['Actuals'].id, include='format')

    if folderId in wbsData:
        print(f"Updating Actuals For Project {fdata['folder_meta'].name} : {folderId}")
        actuals_sheet.update_actuals(client=client, sheet=fdata['sheets']['Actuals'], wbsData=wbsData[folderId])


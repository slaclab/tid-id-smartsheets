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
from . import navigate
import yaml

Columns = ['Sheet Name',  # 1
           'Primary',  # 2
           'Start',  # 3
           'End',  # 4
           'Planned Labor Hours From Budget',  # 5
           'Actual Labor Hours',  # 6
           '% Complete',  # 7
           'Reported Labor Hours',  # 8
           'Reported % Complete',  # 9
           'Estimated Finish',  # 10
           'Status Date',  # 11
           'Notes',  # 12
           'PA Number',  # 13
           'Assigned To',  # 14
           'Baseline Start',  # 15
           'Baseline Finish']  # 16

def check_structure(*, sheet):

    # Check column count
    if len(sheet.columns) != len(Columns):
        print(f"   Wrong number of columns in resource file: Got {len(sheet.columns)} Expect {len(Columns)}.")
        return False

    else:
        ret = True

        for i,v in enumerate(Columns):
            if sheet.columns[i].title != v:
                print(f"   Mismatch resource column name for col {i+1}. Got {sheet.columns[i].title}. Expect {v}.")
                ret = False

        return ret


def check_linked_projects(*, sheet, alist):
    ret = True

    # First create list of
    lsheets = [s.id for s in sheet.scope.sheets]

    for k, v in alist.items():
        if v['Schedule'].id not in lsheets:
            print(f"    Resource file is missing link to {v['name']}")
            ret = False


def check_resource_file(*, client, sheet, alist):
    ret = True

    s = client.Reports.get_report(sheet.id, include='sourceSheets,scope,source')
    print(f"Checking resource file for {s.name}:")

    if not check_structure(sheet=s):
        ret = False
    elif not check_linked_projects(sheet=s, alist=alist):
        ret = False


def check_resource_files(*, client, alist=None, folderId=navigate.TID_RESOURCE_FOLDER):
    folder = client.Folders.get_folder(folderId)

    if alist is None:
        alist = navigate.get_active_list(client=client)

    for s in folder.reports:
        check_resource_file(client=client, sheet=s, alist=alist)

    for f in folder.folders:
        check_resource_files(client=client, alist=alist, folderId=f.id)


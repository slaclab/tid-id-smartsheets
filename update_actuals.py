#-----------------------------------------------------------------------------
# Title      : Update Actuals Script
#-----------------------------------------------------------------------------
# This file is part of the TID ID Smartsheets software platform. It is subject to
# the license terms in the LICENSE.txt file found in the top-level directory
# of this distribution and at:
#    https://confluence.slac.stanford.edu/display/ppareg/LICENSE.html.
# No part of the TID ID Smartsheets software platform, including this file, may be
# copied, modified, propagated, or distributed except according to the terms
# contained in the LICENSE.txt file.
#-----------------------------------------------------------------------------

import tid_ss_lib_v3.navigate
import tid_ss_lib_v3.configuration
import tid_ss_lib_v3.project_list
import tid_ss_lib_v3.division_actuals
import smartsheet
import os
import argparse
import datetime

# Set the argument parser
parser = argparse.ArgumentParser('Smartsheets Actuals Update')

if 'SMARTSHEETS_API' in os.environ:
    defApi = os.environ['SMARTSHEETS_API']
else:
    defApi = ''

parser.add_argument(
    "--div",
    type     = str,
    action   = 'append',
    required = True,
    choices  = [k for k in tid_ss_lib_v3.configuration.division_list],
    help     = "Division for project tracking."
)

parser.add_argument(
    "--key",
    type     = str,
    required = (defApi == ''),
    default  = defApi,
    help     = "API Key from smartsheets. See https://help.smartsheet.com/articles/2482389-generate-API-key"
)

parser.add_argument(
    "--folder",
    action   = 'append',
    required = True,
    help     = "Folder ID to check/fix, or 'all'. Right click on folder, select properties and copy 'Folder ID'",
)

# Get the arguments
args = parser.parse_args()

client = smartsheet.Smartsheet(args.key)

for k in args.div:
    div = tid_ss_lib_v3.configuration.get_division(client=client, div=k)
    folders = []

    if 'all' in args.folder:
        for p in tid_ss_lib_v3.project_list.get_project_list(client=client, div=div):
            folders.append(p['id'])
    else:
        folders = args.folder

    print(f"\n-------------- {k} ------------------------\n")

    print("Collecting Division Actuals")
    data = tid_ss_lib_v3.division_actuals.get_wbs_actuals(client=client,div=div)

    print("Updating Project Actuals Files")
    for p in folders:
        tid_ss_lib_v3.navigate.update_project_actuals(client=client,div=div, folderId=p, wbsData=data)


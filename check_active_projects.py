#-----------------------------------------------------------------------------
# Title      : Check Project Script
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
import tid_ss_lib_v3.division_resource
import smartsheet
import os
import argparse
import datetime

# Set the argument parser
parser = argparse.ArgumentParser('Smartsheets Project Check & Fix')

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
    "--fix",
    action   = 'store_true',
    required = False,
    default  = False,
    help     = "Use to enable fixing of project files.",
)

parser.add_argument(
    "--doCost",
    type     = str,
    required = False,
    choices  = ['None', 'Baseline', 'Contingency'],
    default  = "None",
    help     = "Do cost enable / mode.",
)

parser.add_argument(
    "--doTask",
    action   = 'store_true',
    required = False,
    default  = False,
    help     = "Flag to check for empty task assignments & long duration tasks.",
)

parser.add_argument(
    "--backup",
    type     = str,
    required = False,
    default  = None,
    help     = "Pass backup director to generate a backup"
)

parser.add_argument(
    "--resources",
    type     = str,
    default  = [],
    action   = 'append',
    required = False,
    choices  = ['PD0', 'PD1', 'PD2', 'PD3'],
    help     = "Projects states for tracking."
)

# Get the arguments
args = parser.parse_args()

client = smartsheet.Smartsheet(args.key)

# Generate download directory name
if args.backup is not None:
    backupDir = f"{args.backup}/" + datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    try:
        os.mkdir(backupDir)
    except FileExistsError:
        pass
else:
    backupDir = None

resourceTable={}

for k in args.div:

    div = tid_ss_lib_v3.configuration.get_division(client=client, div=k)

    print(f"\n-------------- {k} ------------------------\n")

    if backupDir is not None:
        doDownload = f"{backupDir}/{div.key}"
        try:
            os.mkdir(doDownload)
        except FileExistsError:
            pass

    else:
        doDownload = False

    tid_ss_lib_v3.navigate.check_project(client=client,
                                         div=div,
                                         folderId=int(div.template_folder),
                                         doFixes=args.fix,
                                         doCost=args.doCost,
                                         doDownload=doDownload,
                                         doTask=args.doTask,
                                         resourceTable=None)

    for p in tid_ss_lib_v3.project_list.get_project_list(client=client, div=div):

        rt = None

        for r in args.resources:
            if r in p['status']:
                rt = resourceTable

        tid_ss_lib_v3.navigate.check_project(client=client,
                                             div=div,
                                             folderId=p['id'],
                                             doFixes=args.fix,
                                             doCost=args.doCost,
                                             doDownload=doDownload,
                                             doTask=args.doTask,
                                             resourceTable=rt)

    if len(resourceTable) > 0:
        tid_ss_lib_v3.division_resource.update(client=client, div=div, resourceTable=resourceTable)


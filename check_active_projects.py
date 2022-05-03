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

#SkipList = [8244071731881860]  # Lsst
SkipList = []

import tid_ss_lib.navigate
import tid_ss_lib.project_sheet
import smartsheet
import os
import argparse

# Set the argument parser
parser = argparse.ArgumentParser('Smartsheets Project Check & Fix')

if 'SMARTSHEETS_API' in os.environ:
    defApi = os.environ['SMARTSHEETS_API']
else:
    defApi = ''

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

# Get the arguments
args = parser.parse_args()

client = smartsheet.Smartsheet(args.key)

for p in tid_ss_lib.project_sheet.get_project_list(client=client):
    if p['id'] in SkipList:
        print(f"Skipping {p['name']}")
    else:
        tid_ss_lib.navigate.check_project(client=client,folderId=p['id'], doFixes=args.fix)

for k in [4013014891423620, # Template
          3142587826628484  # Management
         ]:

    tid_ss_lib.navigate.check_project(client=client,folderId=k, doFixes=True)


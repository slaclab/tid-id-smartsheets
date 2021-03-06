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

import tid_ss_lib.navigate
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
    "--div",
    type     = str,
    required = True,
    default  = False,
    choices  = ['id', 'cds'],
    help     = "Division for project tracking. Either --div=id or --div=cds"
)

parser.add_argument(
    "--fix",
    action   = 'store_true',
    required = False,
    default  = False,
    help     = "Use to enable fixing of project files.",
)

parser.add_argument(
    "--folder",
    action   = 'append',
    required = True,
    help     = "Folder ID to check/fix. Right click on folder, select properties and copy 'Folder ID'",
)

# Get the arguments
args = parser.parse_args()

client = smartsheet.Smartsheet(args.key)
print("")

for p in args.folder:
    tid_ss_lib.navigate.check_project(client=client,div=args.div, folderId=int(p), doFixes=args.fix)


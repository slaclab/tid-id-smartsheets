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

import argparse

# Set the argument parser
parser = argparse.ArgumentParser('Smartsheets Project Check & Fix')

parser.add_argument(
    "--key",
    type     = str,
    required = True,
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
    "--folder",
    action   = 'append',
    required = True,
    help     = "Folder ID to check/fix. Right click on folder, select properties and copy 'Folder ID'",
)

# Get the arguments
args = parser.parse_args()

client = smartsheet.Smartsheet(args.key)

print(f"\nProcessing folder(s) {args.folder} with key {args.key}. Fix = {args.fix}\n")

for p in args.folder:
    tid_ss_lib.navigate.check_project(client=client,folderId=int(p), doFixes=args.fix)


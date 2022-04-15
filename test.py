#-----------------------------------------------------------------------------
# Title      : Test Scripts
#-----------------------------------------------------------------------------
# This file is part of the TID ID Smartsheets software platform. It is subject to
# the license terms in the LICENSE.txt file found in the top-level directory
# of this distribution and at:
#    https://confluence.slac.stanford.edu/display/ppareg/LICENSE.html.
# No part of the TID ID Smartsheets software platform, including this file, may be
# copied, modified, propagated, or distributed except according to the terms
# contained in the LICENSE.txt file.
#-----------------------------------------------------------------------------

import os
import tid_ss_lib.project_sheet
import tid_ss_lib.resource_sheet
import smartsheet  # pip3 install smartsheet-python-sdk

if 'SMARTSHEETS_API' in os.environ:
    api = os.environ['SMARTSHEETS_API']
else:
    api = ''

client = smartsheet.Smartsheet(api)

doFixes = False

#for k in [2771110870706052]:
    #tid_ss_lib.navigate.check_project(client=client,folderId=k, doFixes=doFixes)

tid_ss_lib.resource_sheet.check_resource_files(client=client)

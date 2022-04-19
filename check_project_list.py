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

import tid_ss_lib.project_sheet
import smartsheet  # pip3 install smartsheet-python-sdk
import os

if 'SMARTSHEETS_API' in os.environ:
    api = os.environ['SMARTSHEETS_API']
else:
    import secrets
    api = secrets.API_KEY

doFixes=True

client = smartsheet.Smartsheet(api)

tid_ss_lib.project_sheet.check(client=client, doFixes=doFixes)


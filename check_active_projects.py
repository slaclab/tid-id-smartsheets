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

import secrets
import tid_ss_lib.navigate
import tid_ss_lib.project_sheet
import smartsheet

client = smartsheet.Smartsheet(secrets.API_KEY)

doFixes = False

for p in tid_ss_lib.project_sheet.get_project_list(client=client):
    tid_ss_lib.navigate.check_project(client=client,folderId=p['id'], doFixes=doFixes)

for k in [4013014891423620, # Template
          3142587826628484  # Management
         ]:

    tid_ss_lib.navigate.check_project(client=client,folderId=k, doFixes=doFixes)

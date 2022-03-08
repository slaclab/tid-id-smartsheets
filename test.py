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

import secrets
import tid_ss_lib.navigate
import smartsheet  # pip3 install smartsheet-python-sdk

client = smartsheet.Smartsheet(secrets.API_KEY)

#info = client.Server.server_info()

#print(info)

#exit(1)


if False:
    tid_ss_lib.navigate.check_folders(client=client, doFixes=False)

if True:
    stable = { 4013014891423620: True,   # Template
               3142587826628484: True,   # Management
             }

    for k,v in stable.items():
        tid_ss_lib.navigate.check_project(client=client,folderId=k, doFixes=v)

# Smurf
#tid_ss_lib.navigate.check_project(client=client,folderId=4360781211953028, doFixes=True)
#tid_ss_lib.navigate.check_project(client=client,folderId=6458670872586116, doFixes=True)

#tid_ss_lib.navigate.check_project(client=client,folderId=7748522700236676, doFixes=True)
#tid_ss_lib.navigate.check_project(client=client,folderId=8346297322235780, doFixes=True)
#tid_ss_lib.navigate.check_project(client=client,folderId=2083479090423684, doFixes=True)
#tid_ss_lib.navigate.check_project(client=client,folderId=7919787742390148, doFixes=True)


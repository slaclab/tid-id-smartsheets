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

import tid_ss_lib.secrets
import tid_ss_lib.navigate
import smartsheet

client = smartsheet.Smartsheet(tid_ss_lib.secrets.API_KEY)

#tid_ss_lib.navigate.check_folders(client=client, doFixes=False)

#tid_ss_lib.navigate.check_project(client=client,folderId=1120559233820548, doFixes=False) # UCSC PET Scanner
#tid_ss_lib.navigate.check_project(client=client,folderId=8701540509738884, doFixes=False) # 4D Tracking
#tid_ss_lib.navigate.check_project(client=client,folderId=6454609309919108, doFixes=False) # E-Band Phase 1
#tid_ss_lib.navigate.check_project(client=client,folderId=3024133031257988, doFixes=False) # E-Band Phase 2

tid_ss_lib.navigate.check_project(client=client,folderId=4312801192765316, doFixes=True) # New Magnetron

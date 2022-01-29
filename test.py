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

stable = { 1120559233820548: True,  # UCSC PET
           #8701540509738884: False, # 4D Tracking
           #7854891323418500: False, # Mathusula
           #8002176489416580: False, # Fabulous
           7748522700236676: True,  # RFSOC
           1453019767302020: True,  # LDMX
           6454609309919108: True,  # E-Band Phase 1
           3024133031257988: True,  # E-Band Phase 2
         }

for k,v in stable.items():
    tid_ss_lib.navigate.check_project(client=client,folderId=k, doFixes=v)

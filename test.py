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
import smartsheet

client = smartsheet.Smartsheet(secrets.API_KEY)

if False:
    tid_ss_lib.navigate.check_folders(client=client, doFixes=False)
else:
    stable = { #1120559233820548: True,  # UCSC PET
               #8701540509738884: True,  # 4D Tracking
               #7854891323418500: True,  # Mathusula
               #8002176489416580: True,  # Fabulous
               #7748522700236676: True,  # RFSOC
               #1453019767302020: True,  # LDMX
               #6454609309919108: True,  # E-Band Phase 1
               #3024133031257988: True,  # E-Band Phase 2
               #2771110870706052: True,  # LDMX
               #8244071731881860: True,  # LSST
               #1592107988215684: True,  # Smurf
               #3037956115064708: True # Maps
               #7531399218521988: True # LDRD Frisch
               #2705714624915332: True, # Cryo diamond
               #4260424066590596: True, # FDSOI
               6389267556525956: True, # LGAD
             }

    for k,v in stable.items():
        tid_ss_lib.navigate.check_project(client=client,folderId=k, doFixes=v)

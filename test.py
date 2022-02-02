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

#info = client.Server.server_info()

#print(info)

#exit(1)


if False:
    tid_ss_lib.navigate.check_folders(client=client, doFixes=False)

if True:
    stable = { 2771110870706052: True,  # LDMX
               6454609309919108: True,  # E-Band Phase 1
               3024133031257988: True,  # E-Band Phase 2
               8701540509738884: True,  # 4D Tracking
               #2705714624915332: True,  # Cryo diamond
               #4260424066590596: True,  # FDSOI
               8002176489416580: True,  # Fabulous
               7854891323418500: True,  # Mathusula
               6382760681072516: True,  # HDL Cores
               6389267556525956: True,  # LGAD
               3037956115064708: True,  # Maps
               1453019767302020: True,  # LDMX
               7919787742390148: True,  # LNTPC
               7748522700236676: True,  # RFSOC
               3827405463807876: True,  # Skipper CMOS
               2715125972002692: True,  # LCLS AIP
               7531399218521988: True,  # LDRD Frisch
               8935475567191940: True,  # LDRD Herbst
               8244071731881860: True,  # LSST
               #4312801192765316: True, # Magnetron
               2059817041848196: True,  # NASA
               1592107988215684: True,  # Smurf
               3218618545661828: True,  # Retinal P.
               1120559233820548: True,  # UCSC PET
             }

    for k,v in stable.items():
        tid_ss_lib.navigate.check_project(client=client,folderId=k, doFixes=v)


#tid_ss_lib.navigate.check_project(client=client,folderId=8346297322235780, doFixes=True)
#tid_ss_lib.navigate.check_project(client=client,folderId=2083479090423684, doFixes=True)




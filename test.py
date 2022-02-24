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
    stable = { 4013014891423620: False,  # Template
               2771110870706052: False,  # LDMX
               6454609309919108: False,  # E-Band Phase 1
               3024133031257988: False,  # E-Band Phase 2
               5311293042255748: False,  # Edge ML
               8701540509738884: False,  # 4D Tracking
               2705714624915332: False,  # Cryo diamond
               4260424066590596: False,  # FDSOI
               8002176489416580: False,  # Fabulous
               7854891323418500: False,  # Mathusula
               6382760681072516: False,  # HDL Cores
               6389267556525956: False,  # LGAD
               3037956115064708: False,  # Maps
               1453019767302020: False,  # LDMX
               7919787742390148: False,  # LNTPC
               7748522700236676: False,  # RFSOC
               3827405463807876: False,  # Skipper CMOS
               2715125972002692: False,  # LCLS AIP
               7531399218521988: False,  # LDRD Frisch
               8935475567191940: False,  # LDRD Herbst
               #6826647564380036: False,  # LDRD Rota
               8244071731881860: False,  # LSST
               4312801192765316: False,  # Magnetron  # Errors
               2059817041848196: False,  # NASA
               1592107988215684: False,  # Smurf  # Errors
               3218618545661828: False,  # Retinal P.
               1120559233820548: False,  # UCSC PET
             }

    for k,v in stable.items():
        tid_ss_lib.navigate.check_project(client=client,folderId=k, doFixes=v)


#tid_ss_lib.navigate.check_project(client=client,folderId=8346297322235780, doFixes=True)
#tid_ss_lib.navigate.check_project(client=client,folderId=2083479090423684, doFixes=True)
#tid_ss_lib.navigate.check_project(client=client,folderId=7919787742390148, doFixes=True)


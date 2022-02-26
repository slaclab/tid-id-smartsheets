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

doFix = False

if True:
    stable = { 4013014891423620: True  and doFix,   # Template
               2771110870706052: True  and doFix,   # LDMX
               6454609309919108: True  and doFix,   # E-Band Phase 1
               3024133031257988: True  and doFix,   # E-Band Phase 2
               5311293042255748: True  and doFix,   # Edge ML
               8701540509738884: True  and doFix,   # 4D Tracking
               2705714624915332: True  and doFix,   # Cryo diamond
               4260424066590596: True  and doFix,   # FDSOI
               8002176489416580: True  and doFix,   # Fabulous
               7854891323418500: True  and doFix,   # Mathusula
               6382760681072516: True  and doFix,   # HDL Cores
               6389267556525956: True  and doFix,   # LGAD
               3037956115064708: True  and doFix,   # Maps
               1453019767302020: True  and doFix,   # LDMX
               7919787742390148: True  and doFix,   # LNTPC
               7748522700236676: True  and doFix,   # RFSOC
               3827405463807876: True  and doFix,   # Skipper CMOS
               2715125972002692: True  and doFix,   # LCLS AIP
               7531399218521988: True  and doFix,   # LDRD Frisch
               8935475567191940: True  and doFix,   # LDRD Herbst
               #6826647564380036: False,  # LDRD Rota
               8244071731881860: True  and doFix,  # LSST
               4312801192765316: True  and doFix,   # Magnetron
               2059817041848196: True  and doFix,  # NASA
               1592107988215684: False and doFix,  # Smurf  # Errors need tracking
               3218618545661828: True  and doFix,   # Retinal P.
               1120559233820548: True  and doFix,  # UCSC PET
               8092672926738308: False and doFix,  # epix10K additional, need tracking
               5789524631545732: False and doFix,  # epixm sxr, need tracking
               6335981910550404: False and doFix,  # epix Txi, need tracking
             }

    for k,v in stable.items():
        tid_ss_lib.navigate.check_project(client=client,folderId=k, doFixes=v)


#tid_ss_lib.navigate.check_project(client=client,folderId=7748522700236676, doFixes=True)

#tid_ss_lib.navigate.check_project(client=client,folderId=8346297322235780, doFixes=True)
#tid_ss_lib.navigate.check_project(client=client,folderId=2083479090423684, doFixes=True)
#tid_ss_lib.navigate.check_project(client=client,folderId=7919787742390148, doFixes=True)


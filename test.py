
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

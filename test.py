
import tid_ss_lib.secrets
import tid_ss_lib.navigate
import smartsheet

client = smartsheet.Smartsheet(tid_ss_lib.secrets.API_KEY)
#projects = walk_folders(path='TID/ID',folder_id=TID_ID_ACTIVE_FOLDER)

#for k,v in projects.items():
    #check_project(client=client,folder_id=v)

tid_ss_lib.navigate.check_project(client=client,folderId=1120559233820548)



import smartsheet
import budget_first_row

TID_WORKSPACE          = 4728845933799300
TID_ID_ACTIVE_FOLDER   = 1039693589571460
TID_ID_TEMPLATE_FOLDER = 4013014891423620
TID_ID_LIST_SHEET      = 2931334483076996

TEMPLATE_PREFIX = '[Project] '

def build_template(*, client):
    temp = {}

    folder = client.Folders.get_folder(TID_ID_TEMPLATE_FOLDER)

    for f in folder.sheets:
        name = f.name[len(TEMPLATE_PREFIX):]
        temp[name] = f.id

    return temp


def check_project(*, client, folder_id):
    folder = client.Folders.get_folder(folder_id)

    print(f"Processing project {folder.name}")

    ##########################################################
    # First Make sure folder has all of the neccessary files
    ##########################################################
    tempList = build_template(client=client)
    foundList = {k: None for k in tempList}

    for s in folder.sheets:
        for k in foundList:
            if k in s.name:
                foundList[k] = s

    for k,v in foundList.items():

        # Copy file if it is missing
        if v is None:
            print(f"   Project is missing '{k}' file. Copying.")
            #client.Sheets.copy_sheet(tempList[k], # Source sheet
            #                         smartsheet.models.ContainerDestination({'destination_type': 'folder',
            #                                                                'destination_id': folder.id,
            #                                                                'new_name': folder.name + ' ' + k}))

        # Check for valid naming, rename if need be
        elif not v.name.startswith(folder.name):
            print(f"   Bad sheet name {v.name} Renaming")
            #client.Sheets.update_sheet(v.id, smartsheet.models.Sheet({'name': folder.name + ' ' + k}))


    # First process budget sheet:
    budget   = client.Sheets.get_sheet(foundList['Budget'].id)
    schedule = client.Sheets.get_sheet(foundList['Schedule'].id)

    # Iterate through each row
    budget_first_row.check(row=budget.rows[0])




def walk_folders(*, path, folder_id):
    folder = client.Folders.get_folder(folder_id)
    ret = {}

    path = path + '/' + folder.name

    # No sub folders, this might be a project
    if len(folder.folders) == 0:

        # Skip projects with no sheets
        if len(folder.sheets) != 0:
            ret[path] = folder_id

    else:
        for sub in folder.folders:
            ret.update(walk_folders(path=path, folder_id=sub.id))

    return ret



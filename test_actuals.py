#-----------------------------------------------------------------------------
# Title      : Check Project Script
#-----------------------------------------------------------------------------
# This file is part of the TID ID Smartsheets software platform. It is subject to
# the license terms in the LICENSE.txt file found in the top-level directory
# of this distribution and at:
#    https://confluence.slac.stanford.edu/display/ppareg/LICENSE.html.
# No part of the TID ID Smartsheets software platform, including this file, may be
# copied, modified, propagated, or distributed except according to the terms
# contained in the LICENSE.txt file.
#-----------------------------------------------------------------------------

import tid_ss_lib_v3.division_actuals
import tid_ss_lib_v3.configuration
import tid_ss_lib_v3.navigate
import smartsheet
import os
import argparse

# Set the argument parser
parser = argparse.ArgumentParser('Smartsheets Project Check & Fix')

defApi = os.environ['SMARTSHEETS_API']

client = smartsheet.Smartsheet(defApi)

print("Collecting Division Actuals")
data = tid_ss_lib_v3.division_actuals.get_wbs_actuals(client=client,div=tid_ss_lib_v3.configuration.get_division(client=client, div='id'))

#print(data)

fid = 287585860904836

print()
print(data[fid])

print()
print("Updating Project Actuals")
tid_ss_lib_v3.navigate.update_project_actuals(client=client,div=tid_ss_lib_v3.configuration.get_division(client=client, div='id'), folderId=fid, wbsData=data)


#-----------------------------------------------------------------------------
# Title      : Manipulate Tracking Sheet
#-----------------------------------------------------------------------------
# This file is part of the TID ID Smartsheets software platform. It is subject to
# the license terms in the LICENSE.txt file found in the top-level directory
# of this distribution and at:
#    https://confluence.slac.stanford.edu/display/ppareg/LICENSE.html.
# No part of the TID ID Smartsheets software platform, including this file, may be
# copied, modified, propagated, or distributed except according to the terms
# contained in the LICENSE.txt file.
#-----------------------------------------------------------------------------

# Set formats
#
# https://smartsheet-platform.github.io/api-docs/#formatting
#
# Colors 31 = Dark Blue
#        26 = Dark Gray
#        23 = Blue
#        18 - Gray
#           = White

ToDelete = ['Reporting Variance']

ColData = { 'Status Month':
               { 'position' : 0,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : 'Current Raw Data',
                 'format'   : ",,,,,,,,,18,,,,,,,",
                 'formula'  : None,
                 'link'     : None},

            'Lookup PA' :
               { 'position' : 1,
                 'type'     : 'DATETIME',
                 'forced'   : None,
                 'format'   : None,
                 'formula'  : None,
                 'link'     : None},

            'Monthly Actuals Date' :
               { 'position' : 2,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,13,,,,,",
                 'formula'  : '=VLOOKUP([Lookup PA]@row, {Actuals Range 1}, 5, false)',
                 'link'     : None},

            'Monthly Actuals From Finance' :
               { 'position' : 3,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,13,2,1,2,,",
                 'formula'  : '=VLOOKUP([Lookup PA]@row, {Actuals Range 1}, 3, false)',
                 'link'     : None},

            'Total Actuals From Finance' :
               { 'position' : 4,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,13,2,1,2,,",
                 'formula'  : '=VLOOKUP([Lookup PA]@row, {Actuals Range 1}, 4, false)',
                 'link'     : None},

            'Total Budget' :
               { 'position' : 5,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,13,2,1,2,,",
                 'formula'  : None,
                 'link'     : 'Total Budgeted Cost'},

            'Actual Cost' :
               { 'position' : 6,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,13,2,1,2,,",
                 'formula'  : None,
                 'link'     : 'Actual Cost'},

            'Remaining Funds' :
               { 'position' : 7,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,13,2,1,2,,",
                 'formula'  : None,
                 'link'     : 'Remaining Funds'},

            'Earned Value' :
               { 'position' : 8,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,13,2,1,2,,",
                 'formula'  : None,
                 'link'     : 'Earned Value'},

            'Planned Earned Value' :
               { 'position' : 9,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,13,2,1,2,,",
                 'formula'  : None,
                 'link'     : 'Planned Earned Value'},

            'Cost Variance' :
               { 'position' : 10,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,13,2,1,2,,",
                 'formula'  : None,
                 'link'     : 'Cost Variance'},

            'CPI' :
               { 'position' : 11,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,,2,1,1,,",
                 'formula'  : "=[Earned Value]@row / ([Actual Cost]@row + 0.01)",
                 'link'     : None},

            'Schedule Variance' :
               { 'position' : 12,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,13,2,1,2,,",
                 'formula'  : None,
                 'link'     : 'Schedule Variance'},

            'SPI' :
               { 'position' : 13,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,,2,1,1,,",
                 'formula'  : "=[Earned Value]@row / ([Planned Earned Value]@row + 0.01)",
                 'link'     : None},

            'Reporting Index' :
               { 'position' : 14,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,,2,1,1,,",
                 'formula'  : '=[Total Actuals From Finance]@row / ([Actual Cost]@row + 0.01)',
                 'link'     : None},

            'Tracking Risk' :
               { 'position' : 15,
                 'type'     : 'PICKLIST',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,,,,,,",
                 'formula'  : '=IF(ABS(1 - [Reporting Index]@row) > 0.1, "High", IF(ABS(1 - [Reporting Index]@row) > 0.05, "Medium", "Low"))',
                 'link'     : None},

            'Budget Risk' :
               { 'position' : 16,
                 'type'     : 'PICKLIST',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,,,,,,",
                 'formula'  : '=IF(CPI@row < 0.9, "High", IF(CPI@row < 0.95, "Medium", "Low"))',
                 'link'     : None},

            'Schedule Risk' :
               { 'position' : 17,
                 'type'     : 'PICKLIST',
                 'forced'   : None,
                 'format'   : ",,,,,,,,,18,,,,,,,",
                 'formula'  : '=IF(SPI@row < 0.9, "High", IF(SPI@row < 0.95, "Medium", "Low"))',
                 'link'     : None},

            'Scope Risk' :
               { 'position' : 18,
                 'type'     : 'PICKLIST',
                 'forced'   : None,
                 'format'   : None,
                 'formula'  : None,
                 'link'     : None},

            'Description Of Status' :
               { 'position' : 19,
                 'type'     : 'TEXT_NUMBER',
                 'forced'   : 'My Status Description',
                 'format'   : None,
                 'formula'  : None,
                 'link'     : None},}



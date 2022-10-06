#-----------------------------------------------------------------------------
# Title      : Schedule Sheet Manipulation
#-----------------------------------------------------------------------------
# This file is part of the TID ID Smartsheets software platform. It is subject to
# the license terms in the LICENSE.txt file found in the top-level directory
# of this distribution and at:
#    https://confluence.slac.stanford.edu/display/ppareg/LICENSE.html.
# No part of the TID ID Smartsheets software platform, including this file, may be
# copied, modified, propagated, or distributed except according to the terms
# contained in the LICENSE.txt file.
#-----------------------------------------------------------------------------

import smartsheet
from . import navigate

PlannedPct = '=IF(Start@row > TODAY(), 0, IF(End@row > TODAY(), NETWORKDAYS(Start@row, TODAY()) / Duration@row, 1))'
Contingency = '=IF([Risk Factor]@row = "Low (5% Contingency)", 1.05, IF([Risk Factor]@row = "Medium (10% Contingency)", 1.1, ' \
              'IF([Risk Factor]@row = "Med-High (25% Contingency)", 1.25, IF([Risk Factor]@row = "High (50% Contingency)", 1.5))))'

# https://smartsheet-platform.github.io/api-docs/#formatting
#
# Colors 39 = Dark Blue
#        26 = Dark Gray
#        23 = Blue
#        22 = Green
#        18 - Gray
#           = White

ColData = { 'Task': { 'position'    : 0, 'type': 'TEXT_NUMBER',
                     'top'          : { 'forced': 'Project', 'default': None, 'format': ",3,1,,,,,,2,39,,,,,,1,", 'formula' : None },
                     'ms_top'       : { 'forced': 'M&S',     'default': None, 'format': ",,1,,,,,,2,31,,,,,,1,", 'formula' : None },
                     'ms_parent'    : { 'forced': None,      'default': None, 'format': ",,,,,,,,,23,,,,,,1,", 'formula' : None },
                     'ms_task'      : { 'forced': None,      'default': None, 'format': None,                   'formula' : None },
                     'labor_top'    : { 'forced': 'Labor',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula' : None },
                     'labor_parent' : { 'forced': None,      'default': None, 'format': ",,,,,,,,,23,,,,,,1,", 'formula' : None },
                     'labor_task'   : { 'forced': None,      'default': None, 'format': None,                   'formula' : None }, },

           'Item Notes': { 'position'     : 1, 'type': 'TEXT_NUMBER',
                           'top'          : { 'forced': '',            'default': None, 'format': ",3,1,,,,,,2,39,,,,,,1,", 'formula' : None },
                           'ms_top'       : { 'forced': 'MS_OVERHEAD', 'default': None, 'format': ",,1,,,,,,2,31,,,,,,1,", 'formula' : None },
                           'ms_parent'    : { 'forced': None,          'default': None, 'format': ",,,,,,,,,23,,,,,,1,", 'formula' : None },
                           'ms_task'      : { 'forced': None,          'default': None, 'format': None,                   'formula' : None },
                           'labor_top'    : { 'forced': 'LAB_RATE',    'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula' : None },
                           'labor_parent' : { 'forced': '',            'default': None, 'format': ",,,,,,,,,23,,,,,,1,", 'formula' : None },
                           'labor_task'   : { 'forced': None,          'default': None, 'format': None,                   'formula' : None }, },

           'Budgeted Quantity': { 'position'     : 2, 'type': 'TEXT_NUMBER',
                                  'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())' },
                                  'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,,,,,,", 'formula': '=SUM(CHILDREN())' },
                                  'ms_parent'    : { 'forced': None, 'default': None, 'format': None,                   'formula': '=SUM(CHILDREN())' },
                                  'ms_task'      : { 'forced': None, 'default': '0',  'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                                  'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())' },
                                  'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': '=SUM(CHILDREN())' },
                                  'labor_task'   : { 'forced': None, 'default': '0',  'format': None,                   'formula': None, }, },

           'Cost Per Item': { 'position'     : 3, 'type': 'TEXT_NUMBER',
                              'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                              'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,", 'formula': None },
                              'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                              'ms_task'      : { 'forced': None, 'default': '0',  'format': None,                   'formula': None },
                              'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                              'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                              'labor_task'   : { 'forced': None, 'default': '0',  'format': None,                   'formula': None, }, },

           'Risk Factor': { 'position'     : 4, 'type': 'PICKLIST',
                            'top'          : { 'forced': '',   'default': None,                   'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'ms_top'       : { 'forced': '',   'default': None,                   'format': ",,1,,,,,,2,31,,,,,,,", 'formula': None },
                            'ms_parent'    : { 'forced': '',   'default': None,                   'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                            'ms_task'      : { 'forced': None, 'default': 'Low (5% Contingency)', 'format': None,                   'formula': None },
                            'labor_top'    : { 'forced': '',   'default': None,                   'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'labor_parent' : { 'forced': '',   'default': None,                   'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                            'labor_task'   : { 'forced': None, 'default': 'Low (5% Contingency)', 'format': None,                   'formula': None, }, },

           'Assigned To': { 'position'     : 5, 'type': 'TEXT_NUMBER',
                            'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,", 'formula': None },
                            'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                            'ms_task'      : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                            'labor_task'   : { 'forced': None, 'default': None, 'format': None,                   'formula': None, }, },

           'Predecessors': { 'position'     : 6, 'type': 'PREDECESSOR',
                             'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                             'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,", 'formula': None },
                             'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                             'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                             'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                             'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                             'labor_task'   : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None, }, },

           'Start': { 'position'     : 7, 'type': 'DATETIME',
                      'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,1,1,2,39,,,,,,,", 'formula': None },
                      'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,1,1,2,31,,,,,,,", 'formula': None },
                      'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,1,1,,23,,,,,,,", 'formula': None },
                      'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                      'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                      'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,1,1,,23,,,,,,,", 'formula': None },
                      'labor_task'   : { 'forced': None, 'default': None, 'format': ",,,,,,1,1,,,,,,,,,", 'formula': None, }, },

           'End':   { 'position'     : 8, 'type': 'DATETIME',
                      'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,1,1,2,39,,,,,,,", 'formula': None },
                      'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,1,1,2,31,,,,,,,", 'formula': None },
                      'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,1,1,,23,,,,,,,", 'formula': None },
                      'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                      'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                      'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,1,1,,23,,,,,,,", 'formula': None },
                      'labor_task'   : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None, }, },

           'Duration': { 'position'     : 9, 'type': 'DURATION',
                         'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': '=CALCDURATION(Start1, End1)', },
                         'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,,,,,,", 'formula': '=CALCDURATION(Start2, End2)', },
                         'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,,23,,,,,,,", 'formula': '=CALCDURATION(Start3, End3)', },
                         'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None, },
                         'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())', },
                         'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': '=SUM(CHILDREN())', },
                         'labor_task'   : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None, }, },

           'PA Number': { 'position'     : 10, 'type': 'TEXT_NUMBER',
                          'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                          'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,", 'formula': None },
                          'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                          'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                          'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                          'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                          'labor_task'   : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None, }, },

           'Resource Allocation': { 'position'     : 11, 'type': 'TEXT_NUMBER',
                                    'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,0,1,3,,", 'formula': None },
                                    'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,0,1,3,,", 'formula': None },
                                    'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,0,1,3,,", 'formula': None },
                                    'ms_task'      : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                                    'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                                    'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,0,1,3,,", 'formula': None },
                                    'labor_task'   : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,,,,3,,", 'formula': '=[Budgeted Quantity]@row / (Duration@row * 8)', }, },

           'Total Budgeted Cost':   { 'position'     : 12, 'type': 'TEXT_NUMBER',
                                      'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                      'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                      'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                      'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=[Budgeted Quantity]@row * [Cost Per Item]@row',},
                                      'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                      'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                      'labor_task'   : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,", 'formula': '=[Budgeted Quantity]@row * [Cost Per Item]@row',}, },

           'Actual Labor Hours Or M&S Cost': { 'position'     : 13, 'type': 'TEXT_NUMBER',
                                               'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                               'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                               'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                               'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None,},
                                               'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                               'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,,0,,,,", 'formula': '=SUM(CHILDREN())'},
                                               'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,16,,,0,,,,", 'formula': None, }, },

           '% Complete': { 'position'     : 14, 'type': 'TEXT_NUMBER',
                           'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,,,,3,,", 'formula': None },
                           'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,,0,1,3,1,", 'formula': None },
                           'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                           'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                           'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                           'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,,,,3,,", 'formula': None },
                           'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,16,,,,,3,,", 'formula': None, }, },

           'Actual Cost': { 'position'     : 15, 'type': 'TEXT_NUMBER',
                            'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                            'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                            'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=[Actual Labor Hours Or M&S Cost]@row'},
                            'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                            'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                            'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                            'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,", 'formula': '=[Cost Per Item]@row * [Actual Labor Hours Or M&S Cost]@row',},},

           'Earned Value': { 'position'     : 16, 'type': 'TEXT_NUMBER',
                             'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                             'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                             'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=[Total Budgeted Cost]@row * [% Complete]@row'},
                             'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                             'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                             'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                             'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,", 'formula': '=[Total Budgeted Cost]@row * [% Complete]@row'}, },

           'Planned % Complete': { 'position'     : 17, 'type': 'TEXT_NUMBER',
                                   'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,0,1,3,,", 'formula': None },
                                   'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,0,1,3,,", 'formula': None },
                                   'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': PlannedPct},
                                   'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                                   'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                                   'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,0,1,3,,", 'formula': None },
                                   'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,1,,,,,,2,39,,,,,,,", 'formula': PlannedPct}, },

           'Planned Earned Value': { 'position'     : 18, 'type': 'TEXT_NUMBER',
                                     'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                     'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                     'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=[Planned % Complete]@row * [Total Budgeted Cost]@row'},
                                     'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                     'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                     'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                     'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,", 'formula': '=[Planned % Complete]@row * [Total Budgeted Cost]@row'},},

           'Cost Variance': { 'position'     : 19, 'type': 'TEXT_NUMBER',
                              'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                              'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                              'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=[Earned Value]@row - [Actual Cost]@row'},
                              'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                              'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                              'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                              'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,", 'formula': '=[Earned Value]@row - [Actual Cost]@row'},},

           'Schedule Variance': { 'position'     : 20, 'type': 'TEXT_NUMBER',
                                  'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                  'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                  'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=[Earned Value]@row - [Planned Earned Value]@row'},
                                  'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                  'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                  'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                  'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,", 'formula': '=[Earned Value]@row - [Planned Earned Value]@row'},},

           'Contingency': { 'position'     : 21, 'type': 'TEXT_NUMBER',
                            'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,", 'formula': None },
                            'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': Contingency},
                            'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,", 'formula': None },
                            'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,,0,1,3,,", 'formula': Contingency},},

           'Total Budgeted Cost With Contingency': { 'position'     : 22, 'type': 'TEXT_NUMBER',
                                                     'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                                     'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                                     'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=Contingency@row * [Total Budgeted Cost]@row'},
                                                     'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                                     'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                                     'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                                     'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,", 'formula': '=Contingency@row * [Total Budgeted Cost]@row'},},

           'Earned Value With Contingency': { 'position'     : 23, 'type': 'TEXT_NUMBER',
                                              'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                              'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                              'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=[Total Budgeted Cost With Contingency]@row * [% Complete]@row'},
                                              'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                              'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                              'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                              'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,", 'formula': '=[Total Budgeted Cost With Contingency]@row * [% Complete]@row'},},

           'Cost Variance With Contingency': { 'position'     : 24, 'type': 'TEXT_NUMBER',
                                               'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                               'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                               'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=[Earned Value With Contingency]@row - [Actual Cost]@row'},
                                               'ms_task'      : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                               'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,39,,,,,,,", 'formula': '=SUM(CHILDREN())'},
                                               'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                               'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,", 'formula': '=[Earned Value With Contingency]@row - [Actual Cost]@row'},},
    }


def find_columns(*, client, sheet, doFixes, cData):

    for k,v in cData.items():
        found = False
        for i in range(len(sheet.columns)):
            if sheet.columns[i].title == k:
                found = True
                if v['position'] != i:
                    print(f"Column location mismatch for {k}. Expected at {v['position']}, found at {i}.")
                    v['position'] = i


        if not found:
            print(f"Column not found: {k}.")

            if doFixes is True:
                print(f"Adding column: {k}.")
                col = smartsheet.models.Column({'title': k,
                                                'type': v['type'],
                                                'index': v['position']})

                client.Sheets.add_column(sheet.id, [col])

            return False

    return True

def check_row(*, client, sheet, rowIdx, key, div, cData, doFixes):

    if div == 'id':
        laborRate = navigate.TID_ID_RATE_NOTE
    elif div == 'cds':
        laborRate = navigate.TID_CDS_RATE_NOTE

    row = sheet.rows[rowIdx]

    # Setup row update structure just in case
    new_row = smartsheet.models.Row()
    new_row.id = row.id

    for k,v in cData.items():
        doFormat = False
        idx = v['position']
        data = v[key]

        # First check format
        if data['format'] is not None and row.cells[idx].format != data['format']:
            print(f"   Incorrect format in row {rowIdx+1} {key} cell {idx+1} in project file. Got {row.cells[idx].format} Exp {data['format']}")
            doFormat = True

        # Row has a formula
        if data['formula'] is not None:
            doRow = False

            if not hasattr(row.cells[idx],'formula') or row.cells[idx].formula != data['formula']:
                print(f"   Incorrect formula in row {rowIdx+1} {key} cell {idx+1} in project file. Got {row.cells[idx].formula} Exp {data['formula']}")
                doRow = True

            if doRow or doFormat:
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[idx].id
                new_cell.formula = data['formula']

                if data['format'] is not None:
                    new_cell.format = data['format']

                new_cell.strict = False
                new_row.cells.append(new_cell)

        # Row has a forced value
        elif data['forced'] is not None:
            doRow = False

            if data['forced'] == 'MS_OVERHEAD':
                data['forced'] = navigate.OVERHEAD_NOTE
            elif data['forced'] == 'LB_RATE':
                data['forced'] = laborRate

            if (not (row.cells[idx].value is None and data['forced'] == '')) and (row.cells[idx].value is None or row.cells[idx].value != data['forced']):
                print(f"   Incorrect value in row {rowIdx+1} {key} cell {idx+1} in project file. Got {row.cells[idx].value} Exp {data['forced']}")
                doRow = True

            if doRow or doFormat:
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = sheet.columns[idx].id
                new_cell.value = data['forced']

                if data['format'] is not None:
                    new_cell.format = data['format']

                new_cell.strict = False
                new_row.cells.append(new_cell)

        # Row has a default value and is empty
        elif data['default'] is not None and (row.cells[idx].value is None or row.cells[idx].value == ''):
            print(f"   Missing default value in row {rowIdx+1} {key} cell {idx+1} in project file.")

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[idx].id
            new_cell.value = data['default']

            if data['format'] is not None:
                new_cell.format = data['format']

            new_cell.strict = False
            new_row.cells.append(new_cell)

        # Catch format only case
        elif doFormat:

            new_cell = smartsheet.models.Cell()
            new_cell.column_id = sheet.columns[idx].id
            new_cell.value = row.cells[idx].value
            new_cell.format = data['format']
            new_cell.strict = False
            new_row.cells.append(new_cell)

    if doFixes and len(new_row.cells) != 0:
        print(f"   Applying fixes to project row {rowIdx+1} {key}.")
        client.Sheets.update_rows(sheet.id, [new_row])


def check(*, client, sheet, doFixes, div):
    inLabor = False
    inMS = False
    cData = ColData

    # First check structure
    while True:
        ret = find_columns(client=client, sheet=sheet, doFixes=doFixes, cData=cData)

        if ret is False and doFixes is False:
            return False

        elif ret is True:
            break

        else:
            sheet = client.Sheets.get_sheet(sheet.id, include='format')

    # First do row 0
    check_row(client=client, sheet=sheet, rowIdx=0, key='top', div=div, cData=cData, doFixes=doFixes)

    # First walk through the rows and create a parent list
    parents = set()
    for rowIdx in range(len(sheet.rows)):
        parents.add(sheet.rows[rowIdx].parent_id)

    # Process the rest of the rows
    for rowIdx in range(1,len(sheet.rows)):

        # MS Section
        if sheet.rows[rowIdx].cells[0].value == 'M&S':
            check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='ms_top', div=div, cData=cData, doFixes=doFixes)
            inMS = True
            inLabor = False
        elif sheet.rows[rowIdx].cells[0].value == 'Labor':
            check_row(client=client, sheet=sheet, rowIdx=rowIdx, key='labor_top', div=div, cData=cData, doFixes=doFixes)
            inMS = False
            inLabor = True
        elif inMS or inLabor:

            if inMS:
                key = 'ms'
            else:
                key = 'labor'

            if sheet.rows[rowIdx].id in parents:
                key += '_parent'
            else:
                key += '_task'

            check_row(client=client, sheet=sheet, rowIdx=rowIdx, key=key, div=div, cData=cData, doFixes=doFixes)

    return True
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

PlannedPct = '=IF(Start@row > (DATE(YEAR(TODAY()), MONTH(TODAY()), 1) - 1), 0, ' \
             'IF(End@row > (DATE(YEAR(TODAY()), MONTH(TODAY()), 1) - 1), '\
             'NETWORKDAYS(Start@row, (DATE(YEAR(TODAY()), MONTH(TODAY()), 1) - 1)) / Duration@row, 1))'

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
                     'ms_top'       : { 'forced': 'M&S',     'default': None, 'format': ",,1,,,,,,2,31,,,,,,1,",  'formula' : None },
                     'ms_parent'    : { 'forced': None,      'default': None, 'format': ",,,,,,,,,23,,,,,,1,",    'formula' : None },
                     'ms_task'      : { 'forced': None,      'default': None, 'format': None,                     'formula' : None },
                     'labor_top'    : { 'forced': 'Labor',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,1,",  'formula' : None },
                     'labor_parent' : { 'forced': None,      'default': None, 'format': ",,,,,,,,,23,,,,,,1,",    'formula' : None },
                     'labor_task'   : { 'forced': None,      'default': None, 'format': None,                     'formula' : None }, },

           'Item Notes': { 'position'     : 1, 'type': 'TEXT_NUMBER',
                           'top'          : { 'forced': '',            'default': None, 'format': ",3,1,,,,,,2,39,,,,,,1,", 'formula' : None },
                           'ms_top'       : { 'forced': 'MS_OVERHEAD', 'default': None, 'format': ",,1,,,,,,2,31,,,,,,1,",  'formula' : None },
                           'ms_parent'    : { 'forced': None,          'default': None, 'format': ",,,,,,,,,23,,,,,,1,",    'formula' : None },
                           'ms_task'      : { 'forced': None,          'default': None, 'format': None,                     'formula' : None },
                           'labor_top'    : { 'forced': 'LAB_RATE',    'default': None, 'format': ",,1,,,,,,2,31,,,,,,1,",  'formula' : None },
                           'labor_parent' : { 'forced': '',            'default': None, 'format': ",,,,,,,,,23,,,,,,1,",    'formula' : None },
                           'labor_task'   : { 'forced': None,          'default': None, 'format': None,                     'formula' : None }, },

           'Budgeted Quantity': { 'position'     : 2, 'type': 'TEXT_NUMBER',
                                  'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                                  'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': '=SUM(CHILDREN())' },
                                  'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': '=SUM(CHILDREN())' },
                                  'ms_task'      : { 'forced': None, 'default': '0',  'format': None,                    'formula': None },
                                  'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': '=SUM(CHILDREN())' },
                                  'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': '=SUM(CHILDREN())' },
                                  'labor_task'   : { 'forced': None, 'default': '0',  'format': None,                    'formula': None, }, },

           'Cost Per Item': { 'position'     : 3, 'type': 'TEXT_NUMBER',
                              'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                              'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                              'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                              'ms_task'      : { 'forced': None, 'default': '0',  'format': None,                    'formula': None },
                              'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                              'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                              'labor_task'   : { 'forced': None, 'default': '0',  'format': None,                    'formula': None, }, },

           'Risk Factor': { 'position'     : 4, 'type': 'PICKLIST',
                            'top'          : { 'forced': '',   'default': None,                   'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'ms_top'       : { 'forced': '',   'default': None,                   'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                            'ms_parent'    : { 'forced': '',   'default': None,                   'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                            'ms_task'      : { 'forced': None, 'default': 'Low (5% Contingency)', 'format': None,                    'formula': None },
                            'labor_top'    : { 'forced': '',   'default': None,                   'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                            'labor_parent' : { 'forced': '',   'default': None,                   'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                            'labor_task'   : { 'forced': None, 'default': 'Low (5% Contingency)', 'format': None,                    'formula': None, }, },

           'Assigned To': { 'position'     : 5, 'type': 'TEXT_NUMBER',
                            'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                            'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                            'ms_task'      : { 'forced': '',   'default': None, 'format': ",,,,,,,,,34,,,,,,,",    'formula': None },
                            'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                            'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                            'labor_task'   : { 'forced': None, 'default': None, 'format': None,                    'formula': None, }, },

           'Predecessors': { 'position'     : 6, 'type': 'PREDECESSOR',
                             'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                             'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                             'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                             'ms_task'      : { 'forced': None, 'default': None, 'format': None,                    'formula': None },
                             'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                             'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                             'labor_task'   : { 'forced': None, 'default': None, 'format': None,                    'formula': None, }, },

           'Start': { 'position'     : 7, 'type': 'DATETIME',
                      'top'          : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'ms_top'       : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'ms_parent'    : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'ms_task'      : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'labor_top'    : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'labor_parent' : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'labor_task'   : { 'forced': None, 'default': None, 'format': None, 'formula': None, }, },

           'End':   { 'position'     : 8, 'type': 'DATETIME',
                      'top'          : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'ms_top'       : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'ms_parent'    : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'ms_task'      : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'labor_top'    : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'labor_parent' : { 'forced': None, 'default': None, 'format': None, 'formula': None },
                      'labor_task'   : { 'forced': None, 'default': None, 'format': None, 'formula': None, }, },

           'Duration': { 'position'     : 9, 'type': 'DURATION',
                         'top'          : { 'forced': None, 'default': None, 'format': None, 'formula': None, },
                         'ms_top'       : { 'forced': None, 'default': None, 'format': None, 'formula': None, },
                         'ms_parent'    : { 'forced': None, 'default': None, 'format': None, 'formula': None, },
                         'ms_task'      : { 'forced': None, 'default': None, 'format': None, 'formula': None, },
                         'labor_top'    : { 'forced': None, 'default': None, 'format': None, 'formula': None, },
                         'labor_parent' : { 'forced': None, 'default': None, 'format': None, 'formula': None, },
                         'labor_task'   : { 'forced': None, 'default': None, 'format': None, 'formula': None, }, },

           'PA Number': { 'position'     : 10, 'type': 'TEXT_NUMBER',
                          'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                          'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                          'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                          'ms_task'      : { 'forced': None, 'default': None, 'format': None,                    'formula': None },
                          'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                          'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                          'labor_task'   : { 'forced': None, 'default': None, 'format': None,                    'formula': None, }, },

           'Resource Allocation': { 'position'     : 11, 'type': 'TEXT_NUMBER',
                                    'top'          : { 'forced': '',   'default': None, 'format': None,                       'formula': None },
                                    'ms_top'       : { 'forced': '',   'default': None, 'format': None,                       'formula': None },
                                    'ms_parent'    : { 'forced': '',   'default': None, 'format': None,                       'formula': None },
                                    'ms_task'      : { 'forced': '',   'default': None, 'format': ",,,,,,,,,34,,,0,1,3,,",    'formula': None },
                                    'labor_top'    : { 'forced': '',   'default': None, 'format': None,                       'formula': None },
                                    'labor_parent' : { 'forced': '',   'default': None, 'format': None,                       'formula': None },
                                    'labor_task'   : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,,,,3,,",      'formula': '=[Budgeted Quantity]@row / (Duration@row * 8)', }, },

           'Total Budgeted Cost':   { 'position'     : 12, 'type': 'TEXT_NUMBER',
                                      'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                      'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                      'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                      'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Budgeted Quantity]@row * [Cost Per Item]@row',},
                                      'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                      'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                      'labor_task'   : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Budgeted Quantity]@row * [Cost Per Item]@row',}, },

           'Actual Labor Hours Or M&S Cost': { 'position'     : 13, 'type': 'TEXT_NUMBER',
                                               'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': None, },
                                               'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                               'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                               'ms_task'      : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,16,,13,2,1,2,,",    'formula': None,},
                                               'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,,2,,1,,",      'formula': '=SUM(CHILDREN())'},
                                               'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,,2,,1,,",        'formula': '=SUM(CHILDREN())'},
                                               'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,16,,,2,,1,,",        'formula': None, }, },

           '% Complete': { 'position'     : 14, 'type': 'TEXT_NUMBER',
                           'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,,,,3,,",   'formula': None },
                           'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,,0,1,3,1,", 'formula': None },
                           'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,,0,1,3,1,",   'formula': None },
                           'ms_task'      : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,16,,,0,1,3,1,",   'formula': None },
                           'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,,0,1,3,1,", 'formula': None },
                           'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,,,,3,,",      'formula': None },
                           'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,16,,,,,3,,",      'formula': None, }, },

           'Actual Cost': { 'position'     : 15, 'type': 'TEXT_NUMBER',
                            'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                            'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                            'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                            'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Actual Labor Hours Or M&S Cost]@row'},
                            'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                            'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                            'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Cost Per Item]@row * [Actual Labor Hours Or M&S Cost]@row',},},

           'Remaining Funds': { 'position'     : 16, 'type': 'TEXT_NUMBER',
                                'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Total Budgeted Cost]@row - [Actual Cost]@row'},
                                'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Total Budgeted Cost]@row - [Actual Cost]@row',},},

           'Earned Value': { 'position'     : 17, 'type': 'TEXT_NUMBER',
                             'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                             'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                             'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                             'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Total Budgeted Cost]@row * [% Complete]@row'},
                             'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                             'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                             'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Total Budgeted Cost]@row * [% Complete]@row'}, },

           'Planned % Complete': { 'position'     : 18, 'type': 'TEXT_NUMBER',
                                   'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,0,1,3,,", 'formula': None },
                                   'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,0,1,3,,",  'formula': None },
                                   'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,0,1,3,,",    'formula': None },
                                   'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,,0,1,3,,",    'formula': PlannedPct },
                                   'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,0,1,3,,",  'formula': None },
                                   'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,0,1,3,,",    'formula': None },
                                   'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,,0,1,3,,",    'formula': PlannedPct}, },

           'Planned Earned Value': { 'position'     : 19, 'type': 'TEXT_NUMBER',
                                     'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                     'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                     'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                     'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Planned % Complete]@row * [Total Budgeted Cost]@row'},
                                     'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                     'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                     'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Planned % Complete]@row * [Total Budgeted Cost]@row'},},

           'Cost Variance': { 'position'     : 20, 'type': 'TEXT_NUMBER',
                              'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                              'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                              'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())',},
                              'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Earned Value]@row - [Actual Cost]@row'},
                              'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                              'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                              'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Earned Value]@row - [Actual Cost]@row'},},

           'Schedule Variance': { 'position'     : 21, 'type': 'TEXT_NUMBER',
                                  'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                  'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                  'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                  'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Earned Value]@row - [Planned Earned Value]@row',},
                                  'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                  'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                  'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Earned Value]@row - [Planned Earned Value]@row'},},

           'Contingency': { 'position'     : 22, 'type': 'TEXT_NUMBER',
                            'top'          : { 'forced': '',   'default': None, 'format': ",3,1,,,,,,2,39,,,,,,,", 'formula': None },
                            'ms_top'       : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                            'ms_parent'    : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                            'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,,0,1,3,,", 'formula': Contingency},
                            'labor_top'    : { 'forced': '',   'default': None, 'format': ",,1,,,,,,2,31,,,,,,,",  'formula': None },
                            'labor_parent' : { 'forced': '',   'default': None, 'format': ",,,,,,,,,23,,,,,,,",    'formula': None },
                            'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,,0,1,3,,", 'formula': Contingency},},

           'Total Budgeted Cost With Contingency': { 'position'     : 23, 'type': 'TEXT_NUMBER',
                                                     'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                                     'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                                     'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                                     'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=Contingency@row * [Total Budgeted Cost]@row'},
                                                     'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                                     'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                                     'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=Contingency@row * [Total Budgeted Cost]@row'},},

           'Earned Value With Contingency': { 'position'     : 24, 'type': 'TEXT_NUMBER',
                                              'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                              'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                              'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                              'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Total Budgeted Cost With Contingency]@row * [% Complete]@row'},
                                              'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                              'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",   'formula': '=SUM(CHILDREN())'},
                                              'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,",   'formula': '=[Total Budgeted Cost With Contingency]@row * [% Complete]@row'},},

           'Cost Variance With Contingency': { 'position'     : 25, 'type': 'TEXT_NUMBER',
                                               'top'          : { 'forced': None, 'default': None, 'format': ",3,1,,,,,,2,39,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                               'ms_top'       : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,",  'formula': '=SUM(CHILDREN())'},
                                               'ms_parent'    : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",    'formula': '=SUM(CHILDREN())'},
                                               'ms_task'      : { 'forced': None, 'default': None, 'format': ",,,,,,,,,18,,13,2,1,2,,",    'formula': '=[Earned Value With Contingency]@row - [Actual Cost]@row'},
                                               'labor_top'    : { 'forced': None, 'default': None, 'format': ",,1,,,,,,2,31,,13,2,1,2,,", 'formula': '=SUM(CHILDREN())'},
                                               'labor_parent' : { 'forced': None, 'default': None, 'format': ",,,,,,,,,23,,13,2,1,2,,",   'formula': '=SUM(CHILDREN())'},
                                               'labor_task'   : { 'forced': None, 'default': 0,    'format': ",,,,,,,,,18,,13,2,1,2,,",   'formula': '=[Earned Value With Contingency]@row - [Actual Cost]@row'},},
    }


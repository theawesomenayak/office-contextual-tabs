/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//The following function defines and returns the JSON object that describes the contextual tabs.
//Make changes to the JSON text if you want to modify the contextual tabs.
function getContextualRibbonJSON() {
    const sourceUrl = "https://localhost:5500/";
    return {
        "actions": [
            {
                "id": "idRibbonAction",
                "type": "ExecuteFunction",
                "functionName": "runRibbonAction"
            },
            {
                "id": "showTaskpaneContoso",
                "type": "ShowTaskpane",
                "title": "Contoso task pane",
                "supportPinning": false
            }
        ],
        "tabs": [
            {
                "id": "tabTableData",
                "label": "Spreadsheet Sync",
                "visible": true,
                "groups": [
                    {
                        "id": "grpAccount",
                        "label": "Account",
                        "icon": [
                            {
                                "size": 16,
                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                            },
                            {
                                "size": 32,
                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                            },
                            {
                                "size": 80,
                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                            }
                        ],
                        "controls": [
                            {
                                "type": "Button",
                                "id": "btnSignIn",
                                "enabled": true,
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "label": "Sign In",
                                "toolTip": "Sign In",
                                "superTip": {
                                    "title": "Sign In",
                                    "description": "Sign In"
                                },
                                "actionId": "idRibbonAction"
                            }
                        ]
                    },
                    {
                        "id": "grpSources",
                        "label": "Data Sources",
                        "icon": [
                            {
                                "size": 16,
                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                            },
                            {
                                "size": 32,
                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                            },
                            {
                                "size": 80,
                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                            }
                        ],
                        "controls": [
                            {
                                "type": "Button",
                                "id": "btnAddCompany",
                                "enabled": false,
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "label": "Add Company",
                                "toolTip": "Add Company",
                                "superTip": {
                                    "title": "Add Company",
                                    "description": "Add Company"
                                },
                                "actionId": "idRibbonAction"
                            },
                            {
                                "type": "Menu",
                                "id": "mnuCompanies",
                                "label": "Select Company",
                                "toolTip": "Select Company",
                                "superTip": {
                                    "title": "Select Company",
                                    "description": "Select Company"
                                },
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "items": [
                                    {
                                        "type": "MenuItem",
                                        "id": "itmCo1",
                                        "enabled": false,
                                        "icon": [
                                            {
                                                "size": 16,
                                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                                            },
                                            {
                                                "size": 32,
                                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                                            },
                                            {
                                                "size": 80,
                                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                                            }
                                        ],
                                        "label": "External Excel file",
                                        "toolTip": "Sync with external Excel file",
                                        "superTip": {
                                            "title": "External Excel file",
                                            "description": "Sync with external Excel file"
                                        },
                                        "actionId": "idRibbonAction"
                                    },
                                    {
                                        "type": "MenuItem",
                                        "id": "itmCo2",
                                        "enabled": false,
                                        "icon": [
                                            {
                                                "size": 16,
                                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                                            },
                                            {
                                                "size": 32,
                                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                                            },
                                            {
                                                "size": 80,
                                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                                            }
                                        ],
                                        "label": "SQL Database",
                                        "toolTip": "Sync with SQL Database",
                                        "superTip": {
                                            "title": "SQL Database",
                                            "description": "Sync with SQL Database."
                                        },
                                        "actionId": "idRibbonAction"
                                    }
                                ]
                            },
                            {
                                "type": "Menu",
                                "id": "mnuTables",
                                "label": "Select Table",
                                "toolTip": "Select Table",
                                "superTip": {
                                    "title": "Data Sources",
                                    "description": "Select data source to import from."
                                },
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "items": [
                                    {
                                        "type": "MenuItem",
                                        "id": "itmRec1",
                                        "enabled": false,
                                        "icon": [
                                            {
                                                "size": 16,
                                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                                            },
                                            {
                                                "size": 32,
                                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                                            },
                                            {
                                                "size": 80,
                                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                                            }
                                        ],
                                        "label": "External Excel file",
                                        "toolTip": "Sync with external Excel file",
                                        "superTip": {
                                            "title": "External Excel file",
                                            "description": "Sync with external Excel file"
                                        },
                                        "actionId": "idRibbonAction"
                                    },
                                    {
                                        "type": "MenuItem",
                                        "id": "itmRec2",
                                        "enabled": false,
                                        "icon": [
                                            {
                                                "size": 16,
                                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                                            },
                                            {
                                                "size": 32,
                                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                                            },
                                            {
                                                "size": 80,
                                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                                            }
                                        ],
                                        "label": "SQL Database",
                                        "toolTip": "Sync with SQL Database",
                                        "superTip": {
                                            "title": "SQL Database",
                                            "description": "Sync with SQL Database."
                                        },
                                        "actionId": "idRibbonAction"
                                    }
                                ]
                            },
                            {
                                "type": "Menu",
                                "id": "mnuReports",
                                "label": "Select Report",
                                "toolTip": "Select Report",
                                "superTip": {
                                    "title": "Data Sources",
                                    "description": "Select data source to import from."
                                },
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "items": [
                                    {
                                        "type": "MenuItem",
                                        "id": "itmRep1",
                                        "enabled": false,
                                        "icon": [
                                            {
                                                "size": 16,
                                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                                            },
                                            {
                                                "size": 32,
                                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                                            },
                                            {
                                                "size": 80,
                                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                                            }
                                        ],
                                        "label": "External Excel file",
                                        "toolTip": "Sync with external Excel file",
                                        "superTip": {
                                            "title": "External Excel file",
                                            "description": "Sync with external Excel file"
                                        },
                                        "actionId": "idRibbonAction"
                                    },
                                    {
                                        "type": "MenuItem",
                                        "id": "itmRep2",
                                        "enabled": false,
                                        "icon": [
                                            {
                                                "size": 16,
                                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                                            },
                                            {
                                                "size": 32,
                                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                                            },
                                            {
                                                "size": 80,
                                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                                            }
                                        ],
                                        "label": "SQL Database",
                                        "toolTip": "Sync with SQL Database",
                                        "superTip": {
                                            "title": "SQL Database",
                                            "description": "Sync with SQL Database."
                                        },
                                        "actionId": "idRibbonAction"
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "id": "grpData",
                        "label": "Sync data",
                        "icon": [
                            {
                                "size": 16,
                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                            },
                            {
                                "size": 32,
                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                            },
                            {
                                "size": 80,
                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                            }
                        ],
                        "controls": [
                            {
                                "type": "Button",
                                "id": "btnRefresh",
                                "enabled": false,
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "label": "Refresh",
                                "toolTip": "Refresh",
                                "superTip": {
                                    "title": "Refresh",
                                    "description": "Refresh"
                                },
                                "actionId": "idRibbonAction"
                            },
                            {
                                "type": "Button",
                                "id": "btnDownload",
                                "enabled": false,
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "label": "Download from QBO",
                                "toolTip": "Download from QBO",
                                "superTip": {
                                    "title": "Download from QBO",
                                    "description": "Download from QBO"
                                },
                                "actionId": "idRibbonAction"
                            },
                            {
                                "type": "Button",
                                "id": "btnUpload",
                                "enabled": false,
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "label": "Sync to QBO",
                                "toolTip": "Sync to QBO",
                                "superTip": {
                                    "title": "Sync to QBO",
                                    "description": "Sync to QBO"
                                },
                                "actionId": "idRibbonAction"
                            }
                        ]
                    },
                    {
                        "id": "grpLibrary",
                        "label": "Library",
                        "icon": [
                            {
                                "size": 16,
                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                            },
                            {
                                "size": 32,
                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                            },
                            {
                                "size": 80,
                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                            }
                        ],
                        "controls": [
                            {
                                "type": "Button",
                                "id": "btnTemplates",
                                "enabled": false,
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "label": "Template Library",
                                "toolTip": "Template Library",
                                "superTip": {
                                    "title": "Template Library",
                                    "description": "Template Library"
                                },
                                "actionId": "idRibbonAction"
                            }
                        ]
                    },
                    {
                        "id": "grpTaskpane",
                        "label": "Task pane",
                        "icon": [
                            {
                                "size": 16,
                                "sourceLocation": sourceUrl + "assets/icon-16.png"
                            },
                            {
                                "size": 32,
                                "sourceLocation": sourceUrl + "assets/icon-32.png"
                            },
                            {
                                "size": 80,
                                "sourceLocation": sourceUrl + "assets/icon-80.png"
                            }
                        ],
                        "controls": [
                            {
                                "type": "Button",
                                "id": "btnAudit",
                                "enabled": false,
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "label": "Audit Trail",
                                "toolTip": "Show Contoso task pane",
                                "superTip": {
                                    "title": "Show task pane",
                                    "description": "Show Contoso task pane."
                                },
                                "actionId": "showTaskpaneContoso"
                            },
                            {
                                "type": "Button",
                                "id": "btnTraining",
                                "enabled": false,
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "label": "Training Guide",
                                "toolTip": "Show Contoso task pane",
                                "superTip": {
                                    "title": "Show task pane",
                                    "description": "Show Contoso task pane."
                                },
                                "actionId": "showTaskpaneContoso"
                            },
                            {
                                "type": "Button",
                                "id": "btnHelp",
                                "enabled": true,
                                "icon": [
                                    {
                                        "size": 16,
                                        "sourceLocation": sourceUrl + "assets/icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "sourceLocation": sourceUrl + "assets/icon-32.png"
                                    },
                                    {
                                        "size": 80,
                                        "sourceLocation": sourceUrl + "assets/icon-80.png"
                                    }
                                ],
                                "label": "Help",
                                "toolTip": "Show Contoso task pane",
                                "superTip": {
                                    "title": "Show task pane",
                                    "description": "Show Contoso task pane."
                                },
                                "actionId": "showTaskpaneContoso"
                            }
                        ]
                    }
                ]
            }
        ]
    };
}

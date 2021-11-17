/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
// import "../../assets/icon-16.png";
// import "../../assets/icon-32.png";
// import "../../assets/icon-80.png";
// import { getGlobal } from '../commands/commands.js';
// import { createSampleWorkSheet } from '../utilities/utilities.js';


/* global console, document, Office */

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    //Create the contextual tab
    let g = getGlobal(); 
    await Office.ribbon.requestCreateControls(g.contextualTab);
  }
});
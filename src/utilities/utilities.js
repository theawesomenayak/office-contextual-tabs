/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Excel, Office, console */

/**
 * Enables or disables the Refresh and Submit buttons for table data.
 *
 * @param  {boolean} visible true if the buttons should be enabled; otherwise, false.
 */
function toggleLoggedInButtons(visible) {
  let g = getGlobal();
  g.contextualTab.tabs[0].groups[0].controls[0].enabled = visible;

  g.contextualTab.tabs[0].groups[1].controls[0].enabled = !visible;
  for (let index = 1; index <= 3; index++) {
    g.contextualTab.tabs[0].groups[1].controls[index].items[0].enabled = !visible;
    g.contextualTab.tabs[0].groups[1].controls[index].items[1].enabled = !visible;
  }

  g.contextualTab.tabs[0].groups[2].controls[0].enabled = !visible;
  g.contextualTab.tabs[0].groups[2].controls[1].enabled = !visible;
  g.contextualTab.tabs[0].groups[2].controls[2].enabled = !visible;

  g.contextualTab.tabs[0].groups[3].controls[0].enabled = !visible;

  g.contextualTab.tabs[0].groups[4].controls[0].enabled = !visible;
  g.contextualTab.tabs[0].groups[4].controls[1].enabled = !visible;

  Office.ribbon.requestUpdate(g.contextualTab);
}

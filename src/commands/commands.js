/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, self, window, console */

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

/**
 * Handles all ribbon events from the Contoso contextual tab
 * @param  {} event: The event that was raised.
 */
function runRibbonAction(event) {
  switch (event.source.id) {
    case "btnSignIn": runSignInAction();
      break;
    default: console.log('Event ID: ' + event.source.id + ' was sent, but there is no function handler.');
  }
  event.completed();
}

/**
 * Sign In.
 */
function runSignInAction() {
  toggleLoggedInButtons(false);
}

// the add-in command functions need to be available in global scope
// Globals
const g = getGlobal();

g.contextualTab = getContextualRibbonJSON();
g.excelFileMockData = excelFileMockData;
g.sqlMockData = sqlMockData;
g.runRibbonAction = runRibbonAction;
g.mockDataSource = 'sqlMockData';
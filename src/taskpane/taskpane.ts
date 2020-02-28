/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import "zone.js"; // Required for Angular
import { platformBrowserDynamic } from "@angular/platform-browser-dynamic";
import AppModule from "./app/app.module";
/* global console, document, Office */

Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";
  let jsonFormatWithValues = {
    questionCell: "B2",
    question: "What is the sum of following values?",
    choiceOfAnswerRange: "B3:B5",
    valuesToSum: [[3], [5], [7]],
    gradedCell: "B6",
    wrongAnswerColor: "red",
    rightAnswerColor: "green"
  };

  /**
   * localStorage code goes here
   */
  localStorage.setItem('jsonFormatWithValues', JSON.stringify(jsonFormatWithValues));

  // Bootstrap the app
  platformBrowserDynamic()
    .bootstrapModule(AppModule)
    .catch(error => console.error(error));
};

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("extract-acronyms").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    // Load the entire document body
    const body = context.document.body;
    body.load("text");

    await context.sync();

    // Extract unique acronyms using a regular expression
    const text = body.text;
    const acronyms = [...new Set(text.match(/\b[A-Z]{2,}\b/g))]; // Get unique acronyms

    // Create a new document
    const newDoc = context.application.createDocument();
    await context.sync();

    // Create a table with 2 column
    const table = newDoc.body.insertTable(acronyms.length + 1, 2, Word.InsertLocation.start);
    await context.sync();

    // Insert header
    table.getCell(0, 0).value = "Acronyms";
    table.getCell(0, 1).value = "Definition";

    // Insert acronyms into the table
    acronyms.forEach((acronym, index) => {
      table.getCell(index + 1, 0).value = acronym;
    });

    await context.sync();

    // Open the new document
    newDoc.open();
  });
}

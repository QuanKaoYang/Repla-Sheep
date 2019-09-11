/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import UIkit from 'uikit';
import Icons from 'uikit/dist/js/uikit-icons';

// loads the Icon plugin
UIkit.use(Icons);

// components can be called from the imported UIkit reference
// UIkit.notification('Hello world.');

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("repFile").onchange = replaceFromFile;
    document.getElementById("replace").onclick = gotfile;
    document.getElementById("console").innerText = "onReady OK\n"
  }
});

function gotfile(){
  document.getElementById("console").innerText += "OK got a file\n"
}

export async function replaceFromFile(ev) {
  const fr = new FileReader;
  const repFile = ev.target.files[0];
  let counter = 0;
  fr.readAsText(repFile);
  fr.onload = () => {
    const repPairs = [];
    let lines = fr.result.split("\n");
    for (let line in lines) {
      if (line !== "") {
        repPairs.push(line.split("\t"));
      }
    }
    return Word.run(async context => {
      document.getElementById("console").innerText += "Now In word.run\n"
      for (let repPair in repPairs) {
        counter++
        document.getElementById("console").innerText += String(counter);
        document.getElementById("console").innerText += repPair[0]
        // let searchResults = context.document.body.search(repPair[0], { matchCase: vc.matchCase, useWildcard: vc.useWildcard });
        let searchResults = context.document.body.search(repPair[0]);
        searchResults.load(["text", "font"]);
        await context.sync()
        for (var i = 0; i < searchResults.items.length; i++) {
          searchResults.items[i].insertText("[_" + String(counter) + "_]" + repPair[1] + "[_/_]", "Replace")
          searchResults.items[i].font.color = "purple";
          searchResults.items[i].font.highlightColor = "#FFFF00"; //Yellow
        }
      }
    await context.sync()
    })
  }
}

export async function run() {
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(vm, Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
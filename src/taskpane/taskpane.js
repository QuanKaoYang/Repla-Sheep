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
    document.getElementById("console").innerText = "(^^@)";
    if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
      document.getElementById("console").innerText = "(TT@)";
      document.getElementById("app").style.display = "none";
      document.getElementById("non-app").style.display = "flex";
    }
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    document.getElementById("select_repfile").onchange = readFileByClick;
    document.getElementById("drop_zone").ondragover = handleDragOver;
    document.getElementById("drop_zone").ondragenter = handleDragEnter;
    document.getElementById("drop_zone").ondrop = readFileByDrop;
    document.getElementById("from_textarea").onclick = readFromTextarea;
  }
});

function wl(log) {
  document.getElementById("console").innerText = log;
}

function wlp(log) {
  document.getElementById("console").innerText += "\n" + log;
}

function handleDragEnter(ev) {
  ev.stopPropagation();
  ev.preventDefault();
}

function handleDragOver(ev) {
  ev.stopPropagation();
  ev.preventDefault();
}

function readFileByDrop(ev) {
  ev.stopPropagation();
  ev.preventDefault();
  readRepFile(ev.dataTransfer.files[0])
}

function readFileByClick(ev) {
  ev.stopPropagation();
  ev.preventDefault();
  readRepFile(ev.target.files[0]);
}

export async function readRepFile(repFile) {
  const fr = new FileReader;
  const myAnnotator = new RegExp("^[0-9_/\[\\\]]+$");
  if (!repFile.name.endsWith("tsv") && !repFile.name.endsWith("csv")) {
    wl("(x_x＠)MEEE! I only can eat a 'CSV' or 'TSV' File")
    return
  }
  fr.readAsText(repFile);
  fr.onload = () => {
    const repPairs = [];
    let lines = fr.result.replace(/\r?\n/g, "\n").split("\n");
    for (let line of lines) {
      if (line !== "") {
        const eachTerm = line.split("\t");
        if (eachTerm[0] !== "" && eachTerm[1] !== "") {
          if (!myAnnotator.test(eachTerm[0])) {
            repPairs.push(line.split("\t"));
          }
        }
      }
    }
    wl("(^_^＠) YUMMY!");
    executeReplace(repPairs);
  }
}

export async function readFromClipboard(ev) {
  let cdata;
  const myAnnotator = new RegExp("^[0-9_/\[\\\]]+$");
  const repPairs = [];
  if (window.clipboardData) {
    cdata = window.clipboardData.getData("Text").replace(/\r?\n/g, "\n");
  } else {
    cdata = ev.clipboardData;
  }
  if (cdata == "") {
    return;
  }
  const lines = cdata.split("\n");
  for (let line of lines) {
    if (line !== "") {
      const eachTerm = line.split("\t");
      if (eachTerm[0] !== "" && eachTerm[1] !== "") {
        if (!myAnnotator.test(eachTerm[0])) {
          repPairs.push(line.split("\t"));
        }
      }
    }
    repPairs.push(eachTerm);
  }
  wl("(^_^＠) YUMMY!");
  executeReplace(repPairs);
}

export async function readFromTextarea(ev) {
  let reptext = document.getElementById("reptext").value.replace(/\r?\n/g, "\n");
  const myAnnotator = new RegExp("^[0-9_/\[\\\]]+$");
  const repPairs = [];
  if (reptext == "") {
    return;
  }
  const lines = reptext.split("\n");
  for (let line of lines) {
    if (line !== "") {
      let eachTerm = line.split("\t");
      if (eachTerm[0] !== "" && eachTerm[1] !== "") {
        if (!myAnnotator.test(eachTerm[0])) {
          repPairs.push(line.split("\t"));
        }
      }
    }
  }
  wl("(^_^＠) YUMMY!");
  executeReplace(repPairs);
}

export async function executeReplace(repPairs) {
  wlp("execute!")
  let counter = 0;
  return Word.run(async context => {
    wlp("run run★")
    for (let repPair of repPairs) {
      counter++
      wlp(counter)
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

export async function run() {
  wlp("MeMe, World");
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello, World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var myPref;

Office.onReady(info => {
  console.log("logs")
  if (info.host === Office.HostType.Word) {
    if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {
      document.getElementById("non-app").style.display = "none";
      document.getElementById("app").style.display = "block";
    }
    document.getElementById("drop-zone").ondragover = handleDragOver;
    document.getElementById("drop-zone").ondragenter = handleDragEnter;
    document.getElementById("drop-zone").ondrop = readFileByDrop;
    document.getElementById("select-repfile").onchange = readFileByClick;
    document.getElementById("from-textarea").onclick = ev => readFromTextarea(ev);
    document.getElementById("resetPref").onclick = claerPref;
    myPref = loadPref();
    setInitialPref();
    document.getElementById("use-HL").onchange = ev => setPrefVal(ev, "useHL");
    document.getElementById("use-matchcase").onchange = ev => setPrefBool(ev);
    document.getElementById("use-wildcard").onchange = ev => setPrefBool(ev);
    document.getElementById("use-annotator").onchange = ev => setPrefBool(ev);
    document.getElementById("use-previous").onchange = ev => callPrevious(ev);
  }
});

function initialPref() {
  return {
      useHL : "null",
      matchCase: true,
      wildCard: false,
      annotator: false,
      pre1 : "",
      pre2 : "",
      pre3 : "",
      logs: "",
  }
}

function setInitialPref(){
  document.getElementById("use-HL").value = myPref.useHL;
  document.getElementById("use-matchcase").checked = myPref.matchCase;
  document.getElementById("use-wildcard").checked = myPref.wildCard;
  document.getElementById("use-annotator").checked = myPref.annotator;
  // document.getElementById("logs").innerText = myPref.logs;
}

function claerPref(){
  localStorage.removeItem('myPref');
  myPref = initialPref();
  setInitialPref();
}

function loadPref() {
  let inPref
  if (localStorage.getItem("myPref") !== null) {
    inPref = JSON.parse(localStorage.getItem("myPref"))
  } else {
    inPref = initialPref()
  }
  return inPref
}

function setPrefVal(ev, label){
  myPref[label] = ev.target.value
  localStorage.setItem("myPref", JSON.stringify(myPref));
}

function setPrefBool(ev){
  myPref[ev.target.value] = !myPref[ev.target.value]
  localStorage.setItem("myPref", JSON.stringify(myPref));
}

function setPrefPre(texts){
  myPref.pre3 = myPref.pre2;
  myPref.pre2 = myPref.pre1;
  myPref.pre1 = texts;
  localStorage.setItem("myPref", JSON.stringify(myPref));
}

function setPrefLog(obj) {
  myPref.logs += obj
  localStorage.setItem("myPref", JSON.stringify(myPref));
}


// function wl(log) {
//   document.getElementById("console").innerText = log;
// }

// function wlp(log) {
//   document.getElementById("console").innerText += "\n" + log;
// }

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

async function readRepFile(repFile) {
  const fr = new FileReader;
  const myAnnotator = new RegExp("^[0-9_/\[\\\]]+$");
  if (!repFile.name.endsWith("tsv") && !repFile.name.endsWith("csv")) {
    // wl("(x_x＠)MEEE! I only can eat a 'CSV' or 'TSV' File")
    return
  }
  fr.readAsText(repFile);
  fr.onload = () => {
    const repPairs = [];
    let texts = fr.result.replace(/\r?\n/g, "\n");
    if (repFile.name.endsWith(".csv")){
      texts = texts.replace(/,/g, "\t");
    }
    setPrefPre(texts);
    let lines = texts.split("\n");
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
    // wl("(^_^＠) YUMMY!");
    setPrefLog(repPairs);
    executeReplace(repPairs);
  }
}

async function readFromTextarea(ev) {
  ev.stopPropagation();
  ev.preventDefault();
  let reptext = document.getElementById("reptext").value.replace(/\r?\n/g, "\n");
  if (reptext == "") {
    // wl("oh no")
    return;
  }
  const myAnnotator = new RegExp("^[0-9_/\[\\\]]+$");
  const repPairs = [];
  let texts = reptext.replace(/::/g, "\t");
  setPrefPre(texts);
  let lines = texts.split("\n");
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
  // wl("(^_^＠) YUMMY!");
  executeReplace(repPairs);
}

async function executeReplace(repPairs) {
  let sortedPairs = repPairs.sort((a, b) => {
    if (a[0].length > b[0].length) return -1;
    if (a[0].length < b[0].length) return 1;
    return 0;
  })
  return Word.run(async context => {
    for (let i = 0; i < sortedPairs.length; i++) {
      let searchResults = context.document.body.search(sortedPairs[i][0], { matchCase: myPref.matchcase, useWildcard: myPref.wildcard });
      searchResults.load("text");
      await context.sync()
      for (var j = 0; j < searchResults.items.length; j++) {
        searchResults.items[j].insertText("[_" + String(i) + "_]", "Replace")
      }
    }
    await context.sync()
    for (let i = 0; i < sortedPairs.length; i++) {
      let searchResults = context.document.body.search("[_" + String(i) + "_]", { matchCase: myPref.matchcase, useWildcard: myPref.wildcard });
      searchResults.load(["text", "font"]);
      await context.sync()
      if (searchResults.items.length > 0) {
        const afterReplace = myPref.annotator ? "<_" + String(i) + "_>" + sortedPairs[i][1] + "<_/_>" : sortedPairs[i][1];
        const HLcolor = myPref.useHL === "null" ? null : myPref.useHL;
        for (var j = 0; j < searchResults.items.length; j++) {
          searchResults.items[j].insertText(afterReplace, "Replace")
          searchResults.items[j].font.highlightColor = HLcolor;
        }
      }
    }
    await context.sync()
  })
}

function callPrevious(ev){
  document.getElementById("reptext").value = myPref[ev.target.value];
}

export async function run(repPairs) {
  let c = 0
  let l = ""
  for (let t of repPairs) {
    l += String(c)
    l += t[0]
    c++
    l += String(c)
    l += t[1]
    c++
  }
  return Word.run(async context => {
    const paragraph = context.document.body.insertParagraph(l, Word.InsertLocation.end);
    paragraph.font.color = "blue";
    await context.sync();
  });
}

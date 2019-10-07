/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var myPref;
var indev = false

Office.onReady(info => {
  console.log("logs")
  if (info.host === Office.HostType.Word) {
    if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {
      document.getElementById("non-app").style.display = "none";
      document.getElementById("app").style.display = "block";
    }
    document.getElementById("drop-zone").ondragover = handleDragOver;
    document.getElementById("drop-zone").ondragenter = handleDragEnter;
    document.getElementById("drop-zone").ondrop = ev => readFileAndRep(ev);
    document.getElementById("select-repfile").onchange = ev => readFileAndRep(ev);
    document.getElementById("from-textarea").onclick = ev => readFromTextarea(ev);
    document.getElementById("resetPref").onclick = claerPref;
    myPref = loadPref();
    setInitialPref();
    document.getElementById("use-HL").onchange = ev => setPrefVal(ev, "useHL");
    document.getElementById("use-matchcase").onchange = ev => setPrefBool(ev);
    document.getElementById("use-wildcard").onchange = ev => setPrefBool(ev);
    document.getElementById("use-annotator").onchange = ev => setPrefBool(ev);
    document.getElementById("use-previous").onchange = ev => callPrevious(ev);
    document.getElementById("xl-scol").onchange = ev => setPrefVal(ev, "scol");
    document.getElementById("xl-tcol").onchange = ev => setPrefVal(ev, "tcol");
    document.getElementById("tbx-slang").onchange = ev => setPrefVal(ev, "slang");
    document.getElementById("tbx-tlang").onchange = ev => setPrefVal(ev, "tlang");
    if (indev) {
      document.getElementById("console").style.display = "block";
      wlp("read finish");
    }
    
  }
});

function initialPref() {
  return {
      useHL : "null",
      matchCase: true,
      wildCard: false,
      annotator: false,
      scol: "A",
      tcol: "B",
      slang: "ja-jp-jp",
      tlang: "zh-cn",
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
  document.getElementById("xl-scol").value = myPref.scol;
  document.getElementById("xl-tcol").value = myPref.tcol;
  document.getElementById("tbx-slang").value = myPref.slang;
  document.getElementById("tbx-tlang").value = myPref.tlang;
  // document.getElementById("logs").innerText = myPref.logs;
}

function wlp(log) {
  document.getElementById("console").innerText += "\n" + log;
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

function handleDragEnter(ev) {
  ev.stopPropagation();
  ev.preventDefault();
}

function handleDragOver(ev) {
  ev.stopPropagation();
  ev.preventDefault();
}

function readFileAndRep(ev) {
  ev.stopPropagation();
  ev.preventDefault();
  const repFile = ev.target.files[0]
  const extension = repFile.name.substr(repFile.name.length - 4, 4)
  switch (extension) {
    case ".csv":
      readCTSVFile(repFile, ",")
      break;

    case ".tsv":
      readCTSVFile(repFile, "\t")
      break;

    case "xlsx":
      const scol = myPref.scol.toUpperCase();
      const tcol = myPref.tcol.toUpperCase();
      readXLSXFile(repFile, scol, tcol)
      break;

    case ".tbx":
      readTBXFile(repFile, myPref.slang, myPref.tlang)
      break;

    default:
      break;
  }
}

async function readCTSVFile(repFile, delimiter) {
  const fr = new FileReader;
  const myAnnotator = new RegExp("^[0-9_/\[\\\]]+$");
  fr.readAsText(repFile);
  fr.onload = () => {
    const repPairs = [];
    let texts = fr.result.replace(/\r?\n/g, "\n");
    if (delimiter !== "\t") {
      const delim = new RegExp(delimiter, "g")
      texts = texts.replace(delim, "\t")
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
    setPrefLog(repPairs);
    executeReplace(repPairs);
  }
}

async function readFromTextarea(ev) {
  ev.stopPropagation();
  ev.preventDefault();
  let reptext = document.getElementById("reptext").value.replace(/\r?\n/g, "\n");
  if (reptext == "") {
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
  executeReplace(repPairs);
}

async function readXLSXFile(xlsxFile, scol, tcol) {
  const zip = new JSZip();
  const repPairs = [];
  const rangeColA = new RegExp("^" + scol + "[0-9]+$");
  const rangeColB = new RegExp("^" + tcol + "[0-9]+$");
  zip.loadAsync(xlsxFile).then((inzip) => {
    inzip.folder("xl/").file("sharedStrings.xml").async("string").then((xml) => {
      const sharedStrings = (new DOMParser()).parseFromString(xml, "application/xml");
      const allstrings = [];
      const siNodes = sharedStrings.getElementsByTagName("si");
      for (let siNode of siNodes) {
        let textInSiNode = "";
        const tNodes = siNode.getElementsByTagName("t");
        for (let tNode of tNodes) {
          if (tNode.parentNode.localName !== "rPh") {
            textInSiNode += tNode.textContent;
          }
        }
        allstrings.push(textInSiNode);
      }
      inzip.folder("xl/worksheets/").file("sheet1.xml").async("string").then((xml) => {
        const wsx = (new DOMParser()).parseFromString(xml, "application/xml");
        const rowNodes = wsx.getElementsByTagName("row");
        for (let rowNode of rowNodes) {
          const cellNodes = rowNode.getElementsByTagName("c");
          let gotA = false;
          let gotB = false;
          let avalue = "";
          let bvalue = "";
          for (let cellNode of cellNodes) {
            if (gotA && gotB) {
              break;
            }
            const cellRange = cellNode.getAttribute("r");
            if (!gotA) {
              if (rangeColA.test(cellRange) && cellNode.getAttribute("t") === "s") {
                avalue = allstrings[Number(cellNode.firstChild.textContent)];
                gotA = true;
              }
            } else if (!gotB) {
              if (rangeColB.test(cellRange) && cellNode.getAttribute("t") === "s") {
                bvalue = allstrings[Number(cellNode.firstChild.textContent)];
                gotB = true;
                repPairs.push([avalue, bvalue]);
              }
            }
          }
        }
        let texts = "";
        for (let pair of repPairs) {
          texts += pair.join("\t") + "\n";
        }
        setPrefPre(texts);
        executeReplace(repPairs);
      })
    })
  })
}

async function readTBXFile(repFile, slang, tlang) {
  const repPairs = [];
  const fr = new FileReader;
  fr.readAsText(repFile);
  fr.onload = () => {
    const xml = fr.result;
    const tbxContents = (new DOMParser()).parseFromString(xml, "application/xml");
    const termEntries = tbxContents.getElementsByTagName("text")[0].getElementsByTagName("body")[0].getElementsByTagName("termEntry");
    let termPair;
    for (let termEntry of termEntries) {
      termPair = ["", ""]
      const langSetNodes = termEntry.getElementsByTagName("langSet")
      for (let langSetNode of langSetNodes) {
        const locale = langSetNode.getAttribute("xml:lang")
        switch (locale) {
          case slang:
            termPair[0] = langSetNode.getElementsByTagName("tig")[0].getElementsByTagName("term")[0].textContent  
            break;
          
          case tlang:
            termPair[1] = langSetNode.getElementsByTagName("tig")[0].getElementsByTagName("term")[0].textContent
            break;
        
          default:
            break;
        }
        if (termPair[0] !== "" && termPair[1] !== "") {
          repPairs.push(termPair);
          break;
        }
      }
    }
    let texts = "";
    for (let pair of repPairs) {
      texts += pair.join("\t") + "\n";
    }
    setPrefPre(texts);
    executeReplace(repPairs);
  }
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

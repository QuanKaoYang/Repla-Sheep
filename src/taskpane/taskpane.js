/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var myPref;
var indev = false;

// 開発環境で状態を表示させるための関数
function wlp(log, overwrite) {
  if (overwrite) {
    document.getElementById("console").innerText = log;
  } else {
    document.getElementById("console").innerText += "\n" + log;
  }
}

// 初期設定
Office.onReady(info => {
  // indevがtrueな時は、状態表示用のフッターを見えるようにする
  if (indev) {
    document.getElementById("console").style.display = "block";
    wlp("read started", true);
  }
  // Wordが対応していない場合、sorryページを返す
  if (info.host === Office.HostType.Word && Office.context.requirements.isSetSupported('WordApi', '1.1')) {
    document.getElementById("non-app").style.display = "none";
    document.getElementById("app").style.display = "block";
    // 置換ペインのイベントリスナーの設定をする
    document.getElementById("drop-zone").ondragover = handleDragOver;
    document.getElementById("drop-zone").ondragenter = handleDragEnter;
    document.getElementById("drop-zone").ondrop = ev => handleDrop(ev);
    document.getElementById("select-repfile").onchange = ev => handleClick(ev);
    document.getElementById("from-textarea").onclick = ev => readFromTextarea(ev);
    document.getElementById("resetPref").onclick = claerPref;
    document.getElementById("show-replace").onclick = showReplace;
    document.getElementById("show-preference").onclick = showPreference;
    document.getElementById("replace-done").onclick = closeProcessing;
    // 環境設定のイベントリスナーを設定する
    document.getElementById("use-HL").onchange = ev => setPrefVal(ev, "useHL");
    document.getElementById("use-matchcase").onchange = ev => setPrefBool(ev);
    document.getElementById("use-wildcard").onchange = ev => setPrefBool(ev);
    document.getElementById("use-annotator").onchange = ev => setPrefBool(ev);
    document.getElementById("use-previous").onchange = ev => callPrevious(ev);
    document.getElementById("xl-scol").onchange = ev => setPrefVal(ev, "scol");
    document.getElementById("xl-tcol").onchange = ev => setPrefVal(ev, "tcol");
    document.getElementById("xl-header").onchange = ev => setPrefVal(ev, "xlhead");
    document.getElementById("tbx-slang").onchange = ev => setPrefVal(ev, "slang");
    document.getElementById("tbx-tlang").onchange = ev => setPrefVal(ev, "tlang");
    // 環境設定の読み込み
    myPref = loadPref();
    setInitialPref();

    if (indev) {
      document.getElementById("console").style.display = "block";
      wlp("read finished", true);
    }
  }
});

// デフォルト設定を呼び出す用の関数
function defaultPref() {
  return {
      useHL : "null",
      matchCase: true,
      wildCard: false,
      annotator: false,
      scol: "A",
      tcol: "B",
      xlhead: "0",
      slang: "ja-jp-jp",
      tlang: "zh-cn",
      pre1 : "",
      pre2 : "",
      pre3 : "",
      logs: "",
  }
}

// 初期設定を画面に反映させるための関数
function setInitialPref(){
  document.getElementById("use-HL").value = myPref.useHL;
  document.getElementById("use-matchcase").checked = myPref.matchCase;
  document.getElementById("use-wildcard").checked = myPref.wildCard;
  document.getElementById("use-annotator").checked = myPref.annotator;
  document.getElementById("xl-scol").value = myPref.scol;
  document.getElementById("xl-tcol").value = myPref.tcol;
  document.getElementById("xl-header").value = myPref.xlhead;
  document.getElementById("tbx-slang").value = myPref.slang;
  document.getElementById("tbx-tlang").value = myPref.tlang;
}

// #環境設定関連
// デフォルト設定に戻すための関数
function claerPref(){
  localStorage.removeItem('myPref');
  myPref = defaultPref();
  setInitialPref();
}

// ローカルストレージから環境設定を呼び出す関数
// ローカルストレージに保存されていない場合はデフォルト設定を呼び出す
function loadPref() {
  let inPref
  if (localStorage.getItem("myPref") !== null) {
    inPref = JSON.parse(localStorage.getItem("myPref"))
  } else {
    inPref = defaultPref()
  }
  return inPref
}

// 環境設定を画面とオブジェクトに反映する関数１ 文字列用
function setPrefVal(ev, label){
  myPref[label] = ev.target.value
  localStorage.setItem("myPref", JSON.stringify(myPref));
}

// 環境設定を画面とオブジェクトに反映する関数２ ブーリアン用
function setPrefBool(ev){
  myPref[ev.target.value] = !myPref[ev.target.value]
  localStorage.setItem("myPref", JSON.stringify(myPref));
}

// 環境設定に置換履歴を保存しておくための関数
function setPrefPre(texts){
  myPref.pre3 = myPref.pre2;
  myPref.pre2 = myPref.pre1;
  myPref.pre1 = texts;
  localStorage.setItem("myPref", JSON.stringify(myPref));
}

// 環境設定にログを記録しておく関数
function setPrefLog(obj) {
  myPref.logs += obj
  localStorage.setItem("myPref", JSON.stringify(myPref));
}

// 環境設定から置換履歴を呼び出す関数
function callPrevious(ev) {
  document.getElementById("reptext").value = myPref[ev.target.value];
}

// #画面表示関連
// 置換ペインを呼び出す
function showReplace() {
  document.getElementById("pref-division").style.display = "none"
  document.getElementById("replacing-division").style.display = "none"
  document.getElementById("replace-division").style.display = "block"
}

// 環境設定ペインを呼び出す
function showPreference() {
  document.getElementById("replace-division").style.display = "none"
  document.getElementById("replacing-division").style.display = "none"
  document.getElementById("pref-division").style.display = "block"
}

// 置換実行中の画面を閉じる
function closeProcessing() {
  document.getElementById("replacing-division").style.display = "none";
  document.getElementById("app").style.display = "block";
}

// #置換実行関連
// ドラッグの処理
function handleDragEnter(ev) {
  ev.stopPropagation();
  ev.preventDefault();
}

// ドラッグの処理
function handleDragOver(ev) {
  ev.stopPropagation();
  ev.preventDefault();
}

// ドロップの処理
function handleDrop(ev) {
  ev.stopPropagation();
  ev.preventDefault();
  readFileAndRep(ev.dataTransfer.files[0])
}

// クリックの処理
function handleClick(ev) {
  ev.stopPropagation();
  ev.preventDefault();
  readFileAndRep(ev.target.files[0])
}

// ドロップまたはクリックから呼び出す関数
// ファイルの拡張子から、読込関数を呼び出す
function readFileAndRep(repFile) {
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
      readXLSXFile(repFile, scol, tcol, Number(myPref.xlhead));
      break;

    case ".tbx":
      readTBXFile(repFile, myPref.slang, myPref.tlang)
      break;

    default:
      break;
  }
}

// CSV/TSVを読み込む関数
// 読込結果をexecuteReplaceに渡す
async function readCTSVFile(repFile, delimiter) {
  const fr = new FileReader;
  const myAnnotator = new RegExp("^[0-9_/\[\\\]]+$");
  fr.readAsText(repFile);
  fr.onload = () => {
    const repPairs = [];
    let texts = fr.result.replace(/\r?\n/g, "\n");
    // CSVの場合はカンマをタブに変換しておく
    if (delimiter !== "\t") {
      const delim = new RegExp(delimiter, "g")
      texts = texts.replace(delim, "\t")
    }
    // 置換履歴にセット
    setPrefPre(texts);
    // 置換実行のための二次元配列をつくる
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
    // 置換の実行
    executeReplace(repPairs, repFile.name);
  }
}

// テキストエリアから読み込む関数
// 読込結果をexecuteReplaceに渡す
async function readFromTextarea(ev) {
  ev.stopPropagation();
  ev.preventDefault();
  let reptext = document.getElementById("reptext").value.replace(/\r?\n/g, "\n");
  if (reptext == "") {
    return;
  }
  const myAnnotator = new RegExp("^[0-9_/\[\\\]]+$");
  const repPairs = [];
  // ::をタブに変換しておく
  let texts = reptext.replace(/::/g, "\t");
  // 置換履歴にセット
  setPrefPre(texts);
  // 置換実行のための二次元配列をつくる
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
  // 置換の実行
  executeReplace(repPairs, "From TextArea");
}

// Excelから読み込む関数
// 読込結果をexecuteReplaceに渡す
async function readXLSXFile(xlsxFile, scol, tcol, header) {
  const zip = new JSZip();
  const repPairs = [];
  // 環境設定から原文列・訳文列を設定
  const rangeColA = new RegExp("^" + scol + "[0-9]+$");
  const rangeColB = new RegExp("^" + tcol + "[0-9]+$");
  // zipの中身を読み込む。sheet1のみ
  zip.loadAsync(xlsxFile).then((inzip) => {
    inzip.folder("xl/").file("sharedStrings.xml").async("string").then((xml) => {
      // sharedStringから文字列を読み込み、allstringsに格納する
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
      // sheet1のファイルを読み込む
      inzip.folder("xl/worksheets/").file("sheet1.xml").async("string").then((xml) => {
        const wsx = (new DOMParser()).parseFromString(xml, "application/xml");
        // rowのノードコレクション
        const rowNodes = wsx.getElementsByTagName("row");
        let counter = -1;
        // rowごとに処理を実行
        // 原文列・訳文列のみ、SharedStringの対応番号を読み込む
        // allstringsと照らし合わせて、二次元配列repPairsに格納していく
        for (let rowNode of rowNodes) {
          counter++;
          if (counter < header) {
            continue;
          }
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
        // 置換履歴にセットするため、タグ・改行区切りの文字列に変換
        let texts = "";
        for (let pair of repPairs) {
          texts += pair.join("\t") + "\n";
        }
        // 置換履歴にセット
        setPrefPre(texts);
        // 置換の実行
        executeReplace(repPairs, xlsxFile.name);
      })
    })
  })
}

// TBXからの読み込み
// 読込結果をexecuteReplaceに渡す
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
    executeReplace(repPairs, repFile.name);
  }
}

// 置換の実行
async function executeReplace(repPairs, filename) {
  // 原文の長さに応じて配列をソートし、sortedPairsをつくる
  let sortedPairs = repPairs.sort((a, b) => {
    if (a[0].length > b[0].length) return -1;
    if (a[0].length < b[0].length) return 1;
    return 0;
  })
  // 実行中画面の準備
  document.getElementById("processing-mess").innerText = "Processing...";
  document.getElementById("processing-file").innerText = filename;
  document.getElementById("replace-done").disabled = "true";
  const termTableBody = document.getElementById("term-bd")
  // 過去の原文・訳文用語を消しておく
  while (termTableBody.rows[0]){
    termTableBody.deleteRow(0);
  } 
  // 原文・訳文用語をテーブルにセット
  for (let repPair of sortedPairs) {
    let row = document.createElement("tr");
    let st = document.createElement("td");
    let tt = document.createElement("td");
    st.textContent = repPair[0];
    tt.textContent = repPair[1];
    row.appendChild(st);
    row.appendChild(tt);
    termTableBody.appendChild(row);
  }
  // 表示の切り替え
  document.getElementById("app").style.display = "none";
  document.getElementById("replacing-division").style.display = "block";
  // 置換の実行。プロミスを作る
  const replacing = new Promise(resolve => {
    Word.run(async context => {
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
      await context.sync();
      resolve();
      })
    })
  // プロミスが解決したら、doneボタンを押せるようにし、Finishedと表示させる
  replacing.then(() => {
    document.getElementById("replace-done").disabled = "";
    document.getElementById("processing-mess").innerText = "Finished!";
  })
}


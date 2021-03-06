# Repla-Sheep 
# Word Add-in  Batch Replacer
## 概要

Word上で使用できるOffice Add-inです。

ExcelやCSV、TBXといった用語集ファイルを、直接ドラッグアンドドロップでWordに適用することができます。

## 使い方　～基本編～

マニフェストファイル（manifest.xml）を配置します。

詳細は[マイクロソフト公式](https://docs.microsoft.com/ja-jp/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)を参照し、共有フォルダ内に配置して設定します（NAS等のファイルサーバーがあれば、その中に配置することもできます）。

Wordファイルを開き、ホームタグにあるRepla-Sheepのアイコンをクリックします。

すると一括置換用のペインが現れますので、一括置換に使用したいファイルをドロップしてください。

## 使い方　～応用編～

1. 直接入力
   真ん中のテキストボックスに、一括置換したい用語を直接入力することができます。
   入力の際は ***原文1::訳文1 (改行) 原文2::訳文2 (改行) 原文3::訳文3 (改行) ...*** のように、用語をコロン2つで区切ってください。
2. 履歴の利用
   Repla-Sheepでは、一括置換を実行した履歴を最高３回まで保存しています。
   テキストボックスの下にあるプルダウンから、何回前の履歴を使用するか選択してください。
   選択した内容は、テキストボックス内に表示されます。
   内容に問題がなければ、 **実行** ボタンを押して一括置換を実行してください。

※履歴はユーザーのローカルストレージに保存されています。サーバー側では情報を一切集めていません。

## 使い方　～追加機能～

- Wordでの置換結果に蛍光マーカーを付けることができます。
- 置換結果には、アノテーションタグをつけることができます。
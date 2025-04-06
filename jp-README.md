# docx2md

Microsoft Word の文書ファイル（拡張子 .docx）をMarkdown ファイルへ変換します。

## インストール

```
pip install docx2md
```

## 利用方法

```
usage: python docx2md [-h] [-m] [-v] [--debug] SRC.docx DST.md

positional arguments:
  SRC.docx        Microsoft Word file to read
  DST.md          Markdown file to write

optional arguments:
  -h, --help      show this help message and exit
  -m, --md_table  use Markdown table notation instead of <table>
  -v, --version   show version
  --debug         for debug
```

## 表について

表は ```<table id="table(n)">``` で出力します。
```id``` は 1から始まる出力順です。

```--md_table``` を指定すると、```|``` を使用して出力しますが、タイトル行の項目は ```#``` 固定です。

```
| # | # | # |
|---|---|---|
|a|b|c|
|d|e|f|
|g|h|i|
```

## 画像について

画像は ```<img id="image(n)">``` で出力します。
```id``` は 1から始まる出力順です。

## 変換例

* source: [docx](example/example.docx)
* result: [Markdown](example/example/README.md)


## 変換可能な要素

* テーブル（結合セルも含む）
* リスト（数字付きも箇条書きとなる）
* 見出し
* 埋め込み画像
* 改ページ（```<div class="break"></div>``` へ変換）
* 段落内改行（```<br>``` へ変換）
* テキストボックス（本文に挿入）

## 変換できないもの（分かっているもののみ）

* 目次
* 文字装飾

## ライセンス

[MIT](LINCENSE)

## 更新履歴

- 1.0.2 パッケージングシステムを pyproject.toml へ変更
# docx2md

Microsoft Word の文書ファイル（拡張子 .docx）をMarkdown ファイルへ変換します。

## インストール

```
pip install docx2md
```

## 利用方法

```
python -m docx2md SRC.docx DST.md
```

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

## ライセンス

MIT


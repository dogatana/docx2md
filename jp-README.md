# docx2md

Microsoft Word の文書ファイル（拡張子 .docx）をMarkdown ファイルへ変換します。

## 1. インストール

```
pip install docx2md
```

## 2. 利用方法

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

## 3. 表について

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

## 4. 画像について

画像は `<img id="image(n)">` で出力します。
`id` は 1から始まる出力順です。

## 5. 変換例

* source: [example/example.docx](example/example.docx)
* result: [example/README.md](example/README.md), example/media/image*


## 6. 変換可能な要素

* テーブル（結合セルも含む）
* リスト（数字付きも箇条書きとなる）
* 見出し
* 埋め込み画像
* 改ページ（```<div class="break"></div>``` へ変換）
* 段落内改行（```<br>``` へ変換）
* テキストボックス（本文に挿入）

## 7. 変換できないもの（分かっているもののみ）

* 目次
* 文字装飾

## 8. API

### 8.1. 関数

- docx2md.do_convert

```
>>> help(docx2md.do_convert)
Help on function do_convert in module docx2md.convert:

do_convert(docx_file: str, target_dir='', use_md_table=False) -> str
    convert docx_file to Markdown text and return it

    Args:
        docx_file(str): a file to parse
        target_dir(str): save images into target_dir/media/ if specified
        use_md_table(bool): use Markdown table notation instead of HTHML
    Returns:
        Markdown text(str)
```

### 8.2. クラス

- docx2md.DocxFile
- docx2md.DocxMedia
- docx2md.Converter

それぞれの利用方法は do_convert の実装を参照してください。

```python
def do_convert(docx_file: str, target_dir="", use_md_table=False)  -> str:
    try:
        docx = DocxFile(docx_file)
        media = DocxMedia(docx)
        if target_dir:
            media.save(target_dir)
        converter = Converter(docx.document(), media, use_md_table)
        return converter.convert()
    except Exception as e:
        return f"Exception: {e}"
```


## 9. ライセンス

[MIT](LINCENSE)

## 10. 更新履歴

- 1.0.5 merge PR #7
- 1.0.4 issue #6 修正
- 1.0.3 API 追加
- 1.0.2 パッケージングシステムを pyproject.toml へ変更
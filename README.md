# docx2md

Converts Microsoft Word document files (.docx extension) to Markdown files.

[Japanese](jp-README.md)

## 1. Install

```
pip install docx2md
```

## 2. How to use

```
usage: docx2md [-h] [-m] [-v] [--debug] SRC.docx DST.md

positional arguments:
  SRC.docx        Microsoft Word file to read
  DST.md          Markdown file to write

optional arguments:
  -h, --help      show this help message and exit
  -m, --md_table  use Markdown table notation instead of <table>
  -v, --version   show version
  --debug         for debug
```

## 3. Tables

A table is output as ```<table id="table(n)">```. ```id``` is the order of output, starting from 1.

If ```--md_table``` is specified, the output will use ```|```, but the title line item will be ```#``` fixed.

```
| # | # | # |
|---|---|---|
|a|b|c|
|d|e|f|
|g|h|i|
```

## 4. Pictures

Images will be output as ```<img id="image(n)">```. 
The ```id``` is output in order starting from 1.


## 5. Examples

* source: [example/example.docx](example/example.docx)
* result: [example/README.me](example/README.md), example/media/*

## 6. Elements that can be converted

* Tables (including merged cells)
* Lists (also with numbers as bullets)
* Headings
* Embedded images
* Page breaks (converted to ```<div class="break"></div>```)
* Line breaks within paragraphs (converted to ```<br>```)
* Text boxes (inserted in the body)

## 7. Elements that cannot be converted (only known ones)

* Table of Contents
* Text decoration (bold and etc...)

## 8. API

### 8.1. function

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

### 8.2. class

- docx2md.DocxFile
- docx2md.DocxMedia
- docx2md.Converter

Refer to the do_convert implementation for the usage of each class.

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

## 9. License

[MIT](LINCENSE)

## 10. Changelog

- 1.0.5 merge PR #7
- 1.0.4 fix issue #6
- 1.0.3 add API
- 1.0.2 change packaging system to pyproject.toml
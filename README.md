# docx2md

Converts Microsoft Word document files (.docx extension) to Markdown files.

## Install

```
pip install docx2md
```

## How to use

```
python -m docx2md SRC.docx DST.md
```

## Elements that can be converted

* Tables (including merged cells)
* Lists (also with numbers as bullets)
* Headings
* Embedded images
* Page breaks (converted to ```<div class="break"></div>```)
* Line breaks within paragraphs (converted to ```<br>```)
* Text boxes (inserted in the body)

## Elements that cannot be converted (only known ones)

* Table of Contents

## License

MIT
docx2md
=======

Converts Microsoft Word document files (.docx extension) to Markdown
files.

Install
-------

::

   pip install docx2md

How to use
----------

::

   usage: python -m docx2md [-h] [-m] [-v] [--debug] SRC.docx DST.md

   positional arguments:
     SRC.docx        Microsoft Word file to read
     DST.md          Markdown file to write

   optional arguments:
     -h, --help      show this help message and exit
     -m, --md_table  use Markdown table notation instead of <table>
     -v, --version   show version
     --debug         for debug

Tables
~~~~~~

A table is output as ``<table id="table(n)">``. ``id`` is the order of
output, starting from 1.

If ``--md_table`` is specified, the output will use ``|``, but the title
line item will be ``#`` fixed.

::

   | # | # | # |
   |---|---|---|
   |a|b|c|
   |d|e|f|
   |g|h|i|

Pictures
~~~~~~~~

Images will be output as ``<img id="image(n)">``. The ``id`` is output
in order starting from 1.

Examples
--------

-  source: `docx <example/example.docx>`__
-  result: `Markdown <example/example/README.md>`__

Elements that can be converted
------------------------------

-  Tables (including merged cells)
-  Lists (also with numbers as bullets)
-  Headings
-  Embedded images
-  Page breaks (converted to ``<div class="break"></div>``)
-  Line breaks within paragraphs (converted to ``<br>``)
-  Text boxes (inserted in the body)

Elements that cannot be converted (only known ones)
---------------------------------------------------

-  Table of Contents
-  Text decoration (bold and etcÅc)

License
-------

MIT

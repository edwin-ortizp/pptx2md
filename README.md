# PPTX2MD

[![Downloads](https://pepy.tech/badge/pptx2md)](https://pepy.tech/project/pptx2md)

A tool to convert PowerPoint PPTX files and text-based PDFs into Markdown.

**Preserved formats:**

* Titles. Custom table of contents with fuzzy matching is supported.
* Lists with arbitrary depth.
* Text with **bold**, _italic_, color and [hyperlink](https://github.com/ssine/pptx2md/blob/master/README.md)
* Pictures. They are extracted into image file and relative path is inserted.
* Tables with merged cells.
* Top-to-bottom then left-to-right block order.

**Supported output:**

* Markdown
* [Tiddlywiki](https://tiddlywiki.com/)'s wikitext
* [Madoko](https://www.madoko.net/)
* [Quarto](https://quarto.org/)

_Please star this repo if you like it!_

## Installation & Usage

### Installation

You need to have _[Python](https://www.python.org/)_ version __3.10__ or later and _pip_ installed on your system.

#### Windows

Install `pipx` if it is not available yet:

```powershell
py -m pip install --user pipx
py -m pipx ensurepath
```

Close and reopen the terminal after `ensurepath`, then install this project from the repository folder:

```powershell
py -m pipx uninstall pptx2md
py -m pipx install --editable .
```

If you do not want to use `pipx`, you can install the editable package with `pip`:

```powershell
py -m pip install --editable .
```

#### macOS / Linux

Install `pipx` if it is not available yet:

```sh
python3 -m pip install --user pipx
python3 -m pipx ensurepath
```

Close and reopen the terminal after `ensurepath`, then install this project from the repository folder:

```sh
python3 -m pipx uninstall pptx2md
python3 -m pipx install --editable .
```

### Usage

Once you have installed it, use `pptx2md [filename]` to convert a PPTX or text-based PDF into Markdown.

The default output filename uses the same name as the source file with the `.md` extension, and any pictures extracted (and inserted into .md) will be placed in `/img/` folder.

```sh
pptx2md Modulo\ 0\ -\ Conceptos\ básicos.pptx --disable-color --enable-slides --disable-image --disable-escaping
```

Convert all PPTX and PDF files in the current folder:

```sh
pptx2md --all --disable-color --enable-slides --disable-image --disable-escaping
```

Output:

```text
Modulo 0 - Conceptos básicos.md
```

__Note:__ older .ppt files are not supported, convert them to the new .pptx version first.

### Desktop interface

Install the optional desktop interface in your environment:

```powershell
python -m pip install --editable ".[gui]"
```

From this project, open the Spanish desktop window with:

```powershell
.\.venv\Scripts\python.exe -m pptx2md.gui
```

You can also use `pptx2md-gui` after activating the environment. On Windows the
interface uses Microsoft Edge WebView2 Runtime. Conversion runs locally; the
existing command line remains available without the `gui` extra.

Select several PPTX/PDF files, choose the output folder, and choose one Markdown
per file or a combined document. Images and presenter notes are optional for
PPTX and disabled initially. Existing results are never overwritten: duplicate
names get extensions and/or numeric suffixes. Images live under `img/` in separate
document folders, with relative links from the Markdown.

The interface shows the current document, processed-file count, and individual
successes/errors. A failed file does not stop the remaining files. Combined
documents contain successful conversions only, in selection order, with source
headings; their contents' headings shift down one level (up to six). If every
conversion fails, no empty combined Markdown is created. Wait for conversion to
finish before closing the window; this version has no cancellation.

Windows integration smoke (real hidden WebView2, frontend and conversion worker;
dialog selections and opening external applications are simulated):

```powershell
python tests/gui_smoke.py
```

### Text PDF support

```powershell
pptx2md "2020-Scrum-Guide-Spanish-Latin-South-American.pdf" -o "output/2020-Scrum-Guide-Spanish-Latin-South-American.md"
```

PDF conversion preserves the extracted text, accents, numbered lists, bullet lists,
and headings inferred from font sizes. Wrapped lines and unfinished paragraphs
across pages are joined. Dot-leader tables of contents become lists without page
numbers. Only isolated numeric footers in the bottom 10% of the page are removed.
The output is continuous, without `Slide` labels. `-t` overrides matching heading
levels using the same custom-title file as PPTX; `--page` selects a physical PDF
page, starting at 1, including the cover.

This first PDF reader does not perform OCR, extract images, reconstruct tables,
or reproduce complex layouts or columns. Password-protected or unreadable files,
and selected pages without extractable text, fail before writing the output.
Font-based heading and paragraph detection is heuristic; review the resulting
Markdown when the source has unusual typography. Source text is never summarized
or translated. Slide-, image-, notes-, and PowerPoint-layout-specific options do
not affect PDF extraction; all existing output formatters remain available.

`--all` processes both extensions, case-insensitively. When source stems collide,
outputs include the input extension, for example `guide.pdf.md` and
`guide.pptx.md`. Remaining filename collisions are rejected before conversion.
`-o` is a directory for batch conversion and a filename for single-file conversion.

Development checks (install the development dependencies first):

```powershell
python -m unittest discover -s tests -v
```

To also check every character of the supplied 16-page Scrum Guide (apart from
whitespace, list bullets, numeric footers, and TOC dot leaders/page numbers):

```powershell
$env:SCRUM_GUIDE_PDF = "C:\path\2020-Scrum-Guide-Spanish-Latin-South-American.pdf"
python -m unittest discover -s tests -v
```

__Upgrade & Remove:__

Windows:

```powershell
py -m pipx reinstall pptx2md

py -m pipx uninstall pptx2md
```

macOS / Linux:

```sh
python3 -m pipx reinstall pptx2md

python3 -m pipx uninstall pptx2md
```

## Custom Titles

By default, this tool parse all the pptx titles into `level 1` markdown titles, in order to get a hierarchical table of contents, provide your predefined title list in a file and provide it with `-t` argument.

This is a sample title file (titles.txt):

```
Heading 1
  Heading 1.1
    Heading 1.1.1
  Heading 1.2
  Heading 1.3
Heading 2
  Heading 2.1
  Heading 2.2
    Heading 2.1.1
    Heading 2.1.2
  Heading 2.3
Heading 3
```

The first line with spaces in the begining is considered a second level heading and the number of spaces is the unit of indents. In this case, `  Heading 1.1` will be outputted as `## Heading 1.1` . As it has two spaces at the begining, 2 is the unit of heading indent, so `    Heading 1.1.1` with 4 spaces will be outputted as `### Heading 1.1.1`. Header texts are matched with fuzzy matching, unmatched pptx titles will be regarded as the deepest header.

Use it with `pptx2md [filename] -t titles.txt`.

## Full Arguments 

* `-t [filename]` provide the title file
* `-o [filename]` path of the output file
* `-i [path]` directory of the extracted pictures
* `--all` convert all PPTX and PDF files in the target folder
* `--image-width [width]` the maximum width of the pictures, in px. **If set, images are put as html img tag.**
* `--disable-image` disable the image extraction
* `--disable-escaping` do not attempt to escape special characters
* `--disable-notes` do not add presenter notes
* `--disable-wmf` keep wmf formatted image untouched (avoid exceptions under linux)
* `--disable-color` disable color tags in HTML
* `--enable-slides` deliniate slides `\n---\n`, this can help if you want to convert pptx slides to markdown slides
* `--try-multi-column` try to detect multi-column slides (very slow)
* `--min-block-size [size]` the minimum number of characters for a text block to be outputted
* `--wiki` / `--mdk` if you happen to be using tiddlywiki or madoko, this argument outputs the corresponding markup language
* `--qmd` outputs to the qmd markup language used for [quarto](https://quarto.org/docs/presentations/revealjs/) powered presentations
* `--page [number]` only convert the specified page
* `--keep-similar-titles` keep similar titles and add "(cont.)" to repeated slide titles

Note: install [wand](https://docs.wand-py.org/en/0.6.12/) for better chance of successfully converting wmf images, if needed.

## Screenshots

```
Data Link Layer Design Issues
  Services Provided to the Network Layer
  Framing
  Error Control & Flow Control
Error Detection and Correction
  Error Correcting Code (ECC)
  Error Detecting Code
Elementary Data Link Protocols
Sliding Window Protocols
  One-Bit Sliding Window Protocol
  Protocol Using Go Back N
  Using Selective Repeat
Performance of Sliding Window Protocols
Example Data Link Protocols
  PPP
```

<img src="https://raw.githubusercontent.com/ssine/image_bed/master/pic1.png" height=550 >

* **Top**: Title list file content.
* **Bottom**: The table of contents generated.

![2](https://raw.githubusercontent.com/ssine/image_bed/master/pic2.png)

* **Left**: Source pptx file.
* **Right**: Generated markdown file (rendered by madoko).


## API Usage

You can also use pptx2md programmatically in your Python code:

```python
from pptx2md import convert, ConversionConfig
from pathlib import Path

# Basic usage
convert(
    ConversionConfig(
        pptx_path=Path('presentation.pptx'),
        output_path=Path('output.md'),
        image_dir=Path('img'),
        disable_notes=True
    )
)
```

The `ConversionConfig` class accepts the same parameters as the command line arguments:

- `pptx_path`: Path to the input PPTX file (required)
- `output_path`: Path for the output markdown file (required)
- `image_dir`: Directory for extracted images (required)
- `title_path`: Path to custom titles file
- `image_width`: Maximum width for images in px
- `disable_image`: Skip image extraction
- `disable_escaping`: Skip escaping special characters
- `disable_notes`: Skip presenter notes
- `disable_wmf`: Skip WMF image conversion
- `disable_color`: Skip color tags in HTML
- `enable_slides`: Add slide delimiters
- `try_multi_column`: Attempt to detect multi-column slides
- `min_block_size`: Minimum text block size
- `wiki`: Output in TiddlyWiki format
- `mdk`: Output in Madoko format
- `qmd`: Output in Quarto format
- `page`: Convert only specified page number
- `keep_similar_titles`: Keep similar titles with "(cont.)" suffix



## Detailed Parse Rules

### Text and Layout Processing
* Text blocks are identified in two ways:
  * Paragraphs marked as "body" placeholders in the slide
  * Text shapes containing more than the minimum block size (configurable)
* Lists are generated when paragraphs in a block have different indentation levels
* Single-level paragraphs are output as regular text blocks
* Multi-column layouts can be detected with `--try-multi-column` flag
* Grouped shapes are recursively flattened to process their contents
* Shapes are processed in top-to-bottom, left-to-right order

### Title Handling
* When using custom titles:
  * Fuzzy matching is used to match slide titles with the provided title list
  * Matching score must be > 92 for a match to be accepted
  * Unmatched titles default to the deepest header level
* Similar titles (matching score > 92) are omitted by default unless `--keep-similar-titles` is used

### Formatting and Styling
* Text formatting is preserved through markdown syntax:
  * Bold text from PPT is converted to `**bold**`
  * Italic text is converted to `_italic_`
  * Hyperlinks are preserved as `[text](url)`
* Color handling:
  * Theme colors marked as "Accent 1-6" are preserved
  * RGB colors are converted to HTML color codes
  * Dark theme colors are converted to bold text
  * Color tags can be disabled with `--disable-color`

### Special Elements
* Images:
  * Extracted to specified image directory
  * WMF images are converted to PNG when possible
  * Image width can be constrained with `--image-width`
  * HTML img tags are used when width is specified
* Tables:
  * Merged cells are supported
  * Complex formatting within cells is preserved
* Special characters are escaped by default (can be disabled with `--disable-escaping`)
* Presenter notes are included unless disabled with `--disable-notes`

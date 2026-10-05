"""Read text PDFs into the existing document model; no OCR or layout rendering."""

from collections import Counter
from dataclasses import dataclass
import re

import pdfplumber
from pdfminer.pdfdocument import PDFPasswordIncorrect
from rapidfuzz import process

from pptx2md.types import (GeneralSlide, ListItemElement, ParagraphElement,
                          ParsedPresentation, TextRun, TextStyle, TitleElement)


TOC_ENTRY = re.compile(r'^(.*?)\s*\.{3,}\s*\d+\s*$')
LIST_ENTRY = re.compile(r'^([●•▪◦‣*\-]|\d{1,3}[.)])\s+(.+)$')


@dataclass
class Line:
    text: str
    size: float
    left: float
    top: float
    bottom: float
    page: int
    toc: bool = False
    cover: bool = False


def _runs(text):
    return [TextRun(text=text, style=TextStyle())]


def _append_text(element, text):
    # Keep source hyphens: joining across a physical line must not alter words.
    old = element.content[0].text
    separator = '' if old.endswith(('-', '\u2010', '\u00ad')) else ' '
    element.content[0].text = old + separator + text


def parse_pdf(config):
    if not config.pptx_path.is_file():
        raise FileNotFoundError(config.pptx_path)
    lines = []
    sizes = Counter()
    try:
        with pdfplumber.open(config.pptx_path) as pdf:
            if not pdf.pages:
                raise ValueError('PDF has no pages')
            if config.page is not None and not 1 <= config.page <= len(pdf.pages):
                raise ValueError(f'PDF page must be between 1 and {len(pdf.pages)}')
            for number, page in enumerate(pdf.pages, 1):
                if config.page is not None and number != config.page:
                    continue
                raw = page.extract_text_lines()
                raw = [line for line in raw if not (
                    line['text'].strip().isdigit() and line['top'] > page.height * .90)]
                if not raw:
                    raise ValueError(f'PDF page {number} has no extractable text; OCR is not supported')
                is_toc = sum(bool(TOC_ENTRY.match(line['text'])) for line in raw) >= max(3, len(raw) * .6)
                for line in raw:
                    chars = [char for char in line['chars'] if char['text'].strip()]
                    if not chars:
                        continue
                    size = Counter(round(char['size'], 1) for char in chars).most_common(1)[0][0]
                    text = line['text'].strip()
                    lines.append(Line(text, size, line['x0'], line['top'], line['bottom'], number, is_toc))
                    sizes.update({size: len(chars)})
    except PDFPasswordIncorrect as exc:
        raise ValueError('PDF is password protected; supply an unprotected copy') from exc
    except ValueError:
        raise
    except Exception as exc:
        # pdfplumber may wrap pdfminer's error; inspect the exception chain too.
        cause = exc
        while cause is not None:
            if isinstance(cause, PDFPasswordIncorrect):
                raise ValueError('PDF is password protected; supply an unprotected copy') from exc
            cause = cause.__cause__ or cause.__context__
        raise ValueError(f'could not read PDF {config.pptx_path.name}: {exc}') from exc

    body_size = sizes.most_common(1)[0][0]
    # A sparse opening page with a display title is a cover, not a heading
    # hierarchy containing the authors and publication date.
    first = [line for line in lines if line.page == 1]
    has_cover = bool(first and len(first) <= 8 and max(line.size for line in first) > body_size * 2)
    if has_cover:
        for line in first:
            line.cover = True
    heading_sizes = sorted({line.size for line in lines if not line.cover and not line.toc
                            and line.size > body_size * 1.1}, reverse=True)
    toc_indents = sorted({round(line.left, 0) for line in lines if line.toc})
    toc_titles = {}
    for line in lines:
        match = TOC_ENTRY.match(line.text) if line.toc else None
        if match:
            toc_titles[match[1].strip()] = min(6, toc_indents.index(round(line.left, 0)) + 1 + int(has_cover))

    elements = []
    previous = None
    heading_stack = []
    for line in lines:
        gap = line.top - previous.bottom if previous and previous.page == line.page else None
        toc = TOC_ENTRY.match(line.text) if line.toc else None
        if toc:
            elements.append(ListItemElement(content=_runs(toc[1].strip()),
                                           level=toc_indents.index(round(line.left, 0))))
            previous = line
            continue
        heading = (line.cover and line.size == max(item.size for item in first)) or (
            not line.cover and line.size in heading_sizes)
        if line.text in config.custom_titles:
            heading = True
        if heading:
            while heading_stack and heading_stack[-1][0] <= line.size:
                heading_stack.pop()
            level = min(6, heading_stack[-1][1] + 1) if heading_stack else 1 + int(has_cover and not line.cover)
            # Combine a wrapped heading before looking up its full title.
            if (elements and isinstance(elements[-1], TitleElement) and previous
                    and line.page == previous.page and line.size == previous.size
                    and gap is not None and gap < line.size * .65):
                elements[-1].content += ' ' + line.text
            else:
                elements.append(TitleElement(content=line.text, level=level))
            title = elements[-1]
            known = config.custom_titles or toc_titles
            match = process.extractOne(title.content, known.keys(), score_cutoff=96) if known else None
            if match:
                title.level = known[match[0]]
            heading_stack.append((line.size, title.level))
        else:
            marker = LIST_ENTRY.match(line.text)
            if marker:
                token, text = marker.groups()
                elements.append(ListItemElement(content=_runs(text), level=0,
                                               marker=token if token[0].isdigit() else None))
            else:
                last = elements[-1] if elements else None
                same_page = previous and line.page == previous.page
                continuation = (same_page and gap < body_size * .8) or (
                    previous and not same_page and not re.search(r'[.!?:][\"”\)\]]?$', previous.text))
                # Short field labels (contact details, for example) are separate
                # blocks even when the PDF uses ordinary body line spacing.
                if re.match(r'^[A-ZÁÉÍÓÚÜÑ][^:]{0,39}:(?:\s|$)', line.text):
                    continuation = False
                if (last and isinstance(last, (ParagraphElement, ListItemElement)) and continuation
                        and not line.cover and not previous.cover and not previous.toc
                        and (not isinstance(last, ListItemElement) or line.left >= previous.left)):
                    _append_text(last, line.text)
                else:
                    elements.append(ParagraphElement(content=_runs(line.text)))
        previous = line
    return ParsedPresentation(slides=[GeneralSlide(elements=elements)])

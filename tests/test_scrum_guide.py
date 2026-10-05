"""Optional integration check against the supplied guide, without committing it."""

import os
from pathlib import Path
import re
import tempfile
import unittest

import pdfplumber

from pptx2md import ConversionConfig, convert


@unittest.skipUnless(os.environ.get('SCRUM_GUIDE_PDF'), 'set SCRUM_GUIDE_PDF to the original PDF')
class ScrumGuideTests(unittest.TestCase):
    def test_full_text_preservation(self):
        source = Path(os.environ['SCRUM_GUIDE_PDF'])
        expected = []
        with pdfplumber.open(source) as pdf:
            self.assertEqual(len(pdf.pages), 16)
            for page in pdf.pages:
                for line in page.extract_text_lines():
                    text = line['text'].strip()
                    if text.isdigit() and line['top'] > page.height * .90:
                        continue
                    text = re.sub(r'\s*\.{3,}\s*\d+\s*$', '', text)
                    text = re.sub(r'^[●•▪◦‣*\-]\s+', '', text)
                    expected.append(text)
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'guide.md'
            convert(ConversionConfig(pptx_path=source, output_path=output, image_dir=None))
            text = output.read_text(encoding='utf-8')
        actual = []
        for line in text.splitlines():
            line = re.sub(r'^\s*(?:#{1,6}\s+|\*\s+)', '', line)
            actual.append(re.sub(r'\\([\\`*_{}\[\]<>#!])', r'\1', line))
        normalize = lambda value: re.sub(r'\s+', '', ''.join(value))
        self.assertEqual(normalize(expected), normalize(actual))
        self.assertIn('Creative Commons', text)
        self.assertIn('## Cambios de la Guía Scrum 2017 a la Guía Scrum 2020', text)
        self.assertIn('### Simplificación general del lenguaje para una audiencia más amplia', text)
        self.assertIn('en 1995. Básicamente', text)
        self.assertIn('fueron fundamentales al principio:', text)
        self.assertNotIn('Slide ', text)

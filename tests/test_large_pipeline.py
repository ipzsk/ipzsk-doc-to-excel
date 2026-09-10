import importlib.util
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import large_pdf
from pypdf import PdfWriter

class LargePipelineTests(unittest.TestCase):
    def test_ocr_after_30_pages_and_retry_failed_page(self):
        with tempfile.TemporaryDirectory() as directory:
            w=PdfWriter()
            for _ in range(34): w.add_blank_page(600,800)
            w.write(str(Path(directory)/'source.pdf'))
            db=large_pdf.connect(directory)
            def recognize(**kwargs):
                return ('','OCR failed') if kwargs['page']==32 else ('店铺 A 金额 0012','OCR 待核对')
            with patch.object(large_pdf,'local_ocr',side_effect=recognize):
                s=large_pdf.extract(db,directory,34)
            self.assertEqual(s['extracted'],34);self.assertEqual(s['failed'],1);self.assertEqual(s['total'],33)
            with db: db.execute('DELETE FROM pages WHERE failed=1')
            with patch.object(large_pdf,'local_ocr',return_value=('店铺 A 金额 0012','OCR 待核对')) as ocr:
                s=large_pdf.extract(db,directory,34)
            self.assertEqual(ocr.call_count,1);self.assertEqual(s['failed'],0);self.assertEqual(s['total'],34)
            db.close()

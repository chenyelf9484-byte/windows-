import importlib.util
import io
import unittest
from pathlib import Path


@unittest.skipUnless(importlib.util.find_spec('rapidocr_onnxruntime'), 'optional OCR not installed')
class OCRTests(unittest.TestCase):
    def test_chinese_detection_and_render(self):
        from PIL import Image, ImageDraw, ImageFont
        from docreplace.ocr import OCR, render
        from docreplace.rules import Matcher, Rule
        font = next((p for p in ['C:/Windows/Fonts/msyh.ttc', '/System/Library/Fonts/STHeiti Light.ttc'] if Path(p).exists()), None)
        if not font:
            self.skipTest('no Chinese font')
        im = Image.new('RGB', (900, 180), 'white')
        ImageDraw.Draw(im).text((30, 45), '旧公司名称 ABC 123', font=ImageFont.truetype(font, 48), fill='black')
        buf = io.BytesIO()
        im.save(buf, format='PNG')
        units = OCR().scan('sample.png', buf.getvalue(), Matcher([Rule('旧公司名称', '新企业名称')]))
        self.assertEqual(len(units), 1)
        self.assertFalse(units[0].selected)
        self.assertIn('新企业名称', units[0].after)
        changed = render(buf.getvalue(), units, font)
        Image.open(io.BytesIO(changed)).verify()


if __name__ == '__main__':
    unittest.main()

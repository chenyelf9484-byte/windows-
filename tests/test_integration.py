"""可选的真实文档库往返测试，开发环境有相应库时运行。"""
import importlib.util
import io
import tempfile
import unittest
from pathlib import Path
from docreplace.rules import Matcher, Rule
from docreplace.engine import scan_file, save_file


class IntegrationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.m = Matcher([Rule('旧公司', '新企业')])

    def tearDown(self):
        self.tmp.cleanup()

    @unittest.skipUnless(importlib.util.find_spec('docx'), 'python-docx not installed')
    def test_real_word_roundtrip(self):
        from docx import Document
        d = Document()
        p = d.add_paragraph()
        p.add_run('前旧').bold = True
        p.add_run('公司后').italic = True
        d.add_table(rows=1, cols=1).cell(0, 0).text = '旧公司'
        d.sections[0].header.paragraphs[0].text = '旧公司页眉'
        path = self.root/'word.docx'
        d.save(path)
        r = scan_file(path, self.m)
        self.assertFalse(r.error, r.error)
        self.assertFalse(r.warnings, r.warnings)
        self.assertEqual(len(r.units), 3)
        saved = save_file(r, 'direct')
        self.assertEqual(saved.status, '成功', saved.detail)
        read = Document(path)
        self.assertEqual(read.paragraphs[0].text, '前新企业后')
        self.assertTrue(read.paragraphs[0].runs[0].bold)
        self.assertTrue(read.paragraphs[0].runs[1].italic)
        self.assertEqual(read.tables[0].cell(0, 0).text, '新企业')
        self.assertEqual(read.sections[0].header.paragraphs[0].text, '新企业页眉')

    @unittest.skipUnless(importlib.util.find_spec('openpyxl'), 'openpyxl not installed')
    def test_real_excel_roundtrip(self):
        from openpyxl import Workbook, load_workbook
        from openpyxl.styles import Font
        book = Workbook()
        sheet = book.active
        sheet['A1'] = '旧公司'
        sheet['A1'].font = Font(bold=True)
        sheet['B1'] = '=CONCAT("旧公司",A1)'
        sheet['C1'] = 123
        sheet.merge_cells('A2:C2')
        sheet['A2'] = '旧公司'
        p = self.root/'excel.xlsx'
        book.save(p)
        r = scan_file(p, self.m)
        self.assertFalse(r.error, r.error)
        out = save_file(r, 'direct')
        self.assertEqual(out.status, '成功', out.detail)
        read = load_workbook(p)
        self.assertEqual(read.active['A1'].value, '新企业')
        self.assertTrue(read.active['A1'].font.bold)
        self.assertEqual(read.active['B1'].value, '=CONCAT("旧公司",A1)')
        self.assertEqual(read.active['C1'].value, 123)
        self.assertEqual(read.active['A2'].value, '新企业')
        self.assertIn('A2:C2', [str(v) for v in read.active.merged_cells])
        read.close()

    @unittest.skipUnless(importlib.util.find_spec('pptx'), 'python-pptx not installed')
    def test_real_ppt_roundtrip(self):
        from pptx import Presentation
        from pptx.util import Inches
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(2))
        paragraph = box.text_frame.paragraphs[0]
        paragraph.add_run().text = '旧'
        run = paragraph.add_run()
        run.text = '公司'
        run.font.bold = True
        table = slide.shapes.add_table(1, 1, Inches(1), Inches(4), Inches(6), Inches(1)).table
        table.cell(0, 0).text = '旧公司'
        p = self.root/'ppt.pptx'
        prs.save(p)
        r = scan_file(p, self.m)
        self.assertFalse(r.error, r.error)
        out = save_file(r, 'direct')
        self.assertEqual(out.status, '成功', out.detail)
        read = Presentation(p)
        self.assertEqual(read.slides[0].shapes[0].text, '新企业')
        self.assertEqual(read.slides[0].shapes[1].table.cell(0, 0).text, '新企业')

    @unittest.skipUnless(importlib.util.find_spec('PIL'), 'Pillow not installed')
    def test_picture_render_and_default_unselected(self):
        from PIL import Image
        from docreplace.model import Unit
        from docreplace.ocr import render
        font = next((p for p in ['C:/Windows/Fonts/msyh.ttc', '/System/Library/Fonts/PingFang.ttc', '/System/Library/Fonts/STHeiti Light.ttc'] if Path(p).exists()), None)
        if not font:
            self.skipTest('没有可用中文字体')
        original = Image.new('RGB', (300, 100), 'white')
        buf = io.BytesIO()
        original.save(buf, format='PNG')
        unit = Unit('im', 'im', '旧公司', '新企业', 1, 'image', False,
                    {'part': 'image.png', 'box': [[10, 10], [250, 10], [250, 70], [10, 70]]})
        changed = render(buf.getvalue(), [unit], font)
        read = Image.open(io.BytesIO(changed))
        self.assertEqual(read.size, original.size)
        self.assertNotEqual(read.tobytes(), original.convert('RGBA').tobytes())
        self.assertFalse(unit.selected)


if __name__ == '__main__':
    unittest.main()

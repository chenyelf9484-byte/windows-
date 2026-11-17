import io
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from docreplace.engine import enumerate_files, scan_file, save_file
from docreplace.ooxml import X, apply_patches
from docreplace.rules import Matcher, Rule

class RegressionTests(unittest.TestCase):
    def test_empty_filter(self):
        with tempfile.TemporaryDirectory() as d:
            (Path(d)/'a.txt').write_bytes(b'test')
            self.assertEqual(list(enumerate_files([d], extensions=set())), [])
            self.assertEqual(len(list(enumerate_files([d]))), 1)

    def test_excel_phonetics_and_extensions(self):
        for inline in (False, True):
            with self.subTest(inline=inline), tempfile.TemporaryDirectory() as d:
                p=Path(d)/'a.xlsx'
                text='<t>公司</t><rPh sb="0" eb="2"><t>gongsi</t></rPh>'
                payload='<is>'+text+'</is>' if inline else '<v>0</v>'
                sheet=f'<worksheet xmlns="{X}"><sheetData><row><c r="A1" t="'+('inlineStr' if inline else 's')+'">'+payload+'<extLst><ext uri="keep"/></extLst></c></row></sheetData></worksheet>'
                with zipfile.ZipFile(p,'w') as z:
                    z.writestr('xl/worksheets/sheet1.xml',sheet)
                    z.writestr('xl/sharedStrings.xml',f'<sst xmlns="{X}"><si>{text}</si></sst>')
                self.assertEqual(scan_file(p,Matcher([Rule('gongsi','bad')])).units,[])
                result=scan_file(p,Matcher([Rule('公司','新公司')]))
                self.assertEqual(result.units[0].before,'公司')
                saved=save_file(result,'direct');self.assertEqual(saved.status,'成功',saved.detail)
                with zipfile.ZipFile(p) as z:
                    xml=z.read('xl/worksheets/sheet1.xml').decode()
                    self.assertIn('<extLst><ext uri="keep"/></extLst>',xml)
                    self.assertNotIn('gongsi',xml)
                    self.assertIn('新公司',xml)

    def test_patch_bounds(self):
        self.assertEqual(apply_patches('abcdef',[(1,3,'X'),(4,6,'Y')]),'aXdY')
        with self.assertRaises(ValueError): apply_patches('abc',[(0,2,'X'),(1,3,'Y')])

    def test_image_orientation_and_size_guard(self):
        try:
            from PIL import Image
        except ImportError:
            self.skipTest('Pillow unavailable')
        from docreplace.ocr import load_image
        im=Image.new('RGB',(80,40),'white');exif=im.getexif();exif[274]=6
        buf=io.BytesIO();im.save(buf,format='JPEG',exif=exif)
        self.assertEqual(load_image(buf.getvalue()).size,(40,80))
        with patch('docreplace.ocr.MAX_PIXELS',100):
            with self.assertRaises(ValueError): load_image(buf.getvalue())

if __name__=='__main__': unittest.main()

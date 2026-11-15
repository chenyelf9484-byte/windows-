import codecs
import csv
import io
import tempfile
import unittest
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
from docreplace.rules import Matcher, Rule
from docreplace.engine import scan_file, save_file, enumerate_files
from docreplace.ooxml import W, X, A


class EngineTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.matcher = Matcher([Rule('旧公司', '新企业集团'), Rule('旧', '古')])

    def tearDown(self):
        self.tmp.cleanup()

    def package(self, name, parts):
        p = self.root/name
        with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as z:
            for key, value in parts.items():
                z.writestr(key, value)
        return p

    def test_non_cascading_longest_and_literal(self):
        m = Matcher([Rule('甲', '乙'), Rule('乙', '丙'), Rule('甲乙', '$\\')])
        self.assertEqual(m.replace('甲乙 甲 乙'), '$\\ 乙 丙')

    def test_unicode_case_and_duplicate(self):
        self.assertEqual(Matcher([Rule('abc', '好')], True).replace('ABC'), '好')
        with self.assertRaises(ValueError):
            Matcher([Rule('abc', 'a'), Rule('ABC', 'b')], True)

    def test_docx_cross_runs_preserves_other_parts_and_format(self):
        xml = '<w:document xmlns:w="%s"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>前旧</w:t></w:r><w:r><w:t>公司后</w:t></w:r></w:p></w:body></w:document>' % W
        p = self.package('a.docx', {'word/document.xml': xml, 'word/styles.xml': '<styles/>', 'word/media/image1.png': b'unchanged'})
        r = scan_file(p, self.matcher)
        self.assertFalse(r.error)
        self.assertEqual(r.units[0].after, '前新企业集团后')
        result = save_file(r, 'copy', self.root/'out')
        self.assertEqual(result.status, '成功', result.detail)
        with zipfile.ZipFile(result.destination) as z:
            root = ET.fromstring(z.read('word/document.xml'))
            self.assertEqual(''.join(e.text or '' for e in root.iter('{%s}t' % W)), '前新企业集团后')
            self.assertIsNotNone(root.find('.//{%s}b' % W))
            self.assertEqual(z.read('word/media/image1.png'), b'unchanged')
            self.assertEqual(z.read('word/styles.xml'), b'<styles/>')
        with zipfile.ZipFile(p) as z:
            self.assertEqual(z.read('word/document.xml').decode(), xml)

    def test_selective_same_paragraph_and_nested_textbox(self):
        xml = '<w:document xmlns:w="%s"><w:p><w:r><w:t>旧公司</w:t><w:txbxContent><w:p><w:r><w:t>旧公司</w:t></w:r></w:p></w:txbxContent></w:r></w:p></w:document>' % W
        p = self.package('a.docx', {'word/document.xml': xml})
        r = scan_file(p, self.matcher)
        self.assertEqual(len(r.units), 2, r.warnings)
        r.units[1].selected = False
        result = save_file(r, 'direct')
        self.assertEqual(result.status, '成功', result.detail)
        with zipfile.ZipFile(p) as z:
            text = z.read('word/document.xml').decode()
            self.assertEqual(text.count('新企业集团'), 1)
            self.assertEqual(text.count('旧公司'), 1)

    def test_do_not_match_across_tab_or_paragraph(self):
        xml = '<w:document xmlns:w="%s"><w:p><w:r><w:t>旧</w:t><w:tab/><w:t>公司</w:t></w:r></w:p><w:p><w:r><w:t>旧</w:t></w:r></w:p><w:p><w:r><w:t>公司</w:t></w:r></w:p></w:document>' % W
        p = self.package('a.docx', {'word/document.xml': xml})
        r = scan_file(p, Matcher([Rule('旧公司', '新')]))
        self.assertEqual(r.units, [])

    def test_excel_shared_selection_formula_and_number(self):
        sheet = '<worksheet xmlns="%s"><sheetData><row r="1"><c r="A1" t="s" s="2"><v>0</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="str"><f>"旧公司"</f><v>旧公司</v></c><c r="D1"><v>123</v></c></row></sheetData></worksheet>' % X
        shared = '<sst xmlns="%s"><si><r><rPr><b/></rPr><t>旧</t></r><r><t>公司</t></r></si></sst>' % X
        p = self.package('a.xlsx', {'xl/worksheets/sheet1.xml': sheet, 'xl/sharedStrings.xml': shared})
        r = scan_file(p, self.matcher)
        self.assertFalse(r.error, r.error)
        self.assertEqual(len(r.units), 2)
        r.units[1].selected = False
        out = save_file(r, 'direct')
        self.assertEqual(out.status, '成功', out.detail)
        with zipfile.ZipFile(p) as z:
            root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
            cells = list(root.iter('{%s}c' % X))
            self.assertEqual(cells[0].get('t'), 'inlineStr')
            self.assertEqual(cells[0].get('s'), '2')
            self.assertEqual(''.join(e.text or '' for e in cells[0].iter('{%s}t' % X)), '新企业集团')
            self.assertEqual(cells[1].get('t'), 's')
            self.assertEqual(cells[2].find('{%s}f' % X).text, '"旧公司"')
            self.assertEqual(z.read('xl/sharedStrings.xml').decode(), shared)

    def test_excel_inline_prefixed(self):
        xml = '<x:worksheet xmlns:x="%s"><x:sheetData><x:row><x:c r="A1" t="inlineStr"><x:is><x:t>旧公司</x:t></x:is></x:c></x:row></x:sheetData></x:worksheet>' % X
        p = self.package('a.xlsx', {'xl/worksheets/sheet1.xml': xml})
        r = scan_file(p, self.matcher)
        self.assertFalse(r.error, r.error)
        self.assertEqual(save_file(r, 'direct').status, '成功')

    def test_ppt_cross_runs_and_master_switch(self):
        xml = '<a:p xmlns:a="%s"><a:r><a:t>旧</a:t></a:r><a:r><a:t>公司</a:t></a:r></a:p>' % A
        p = self.package('a.pptx', {'ppt/slides/slide1.xml': xml, 'ppt/slideMasters/slideMaster1.xml': xml, 'ppt/other.bin': b'animation'})
        r = scan_file(p, self.matcher)
        self.assertEqual(len(r.units), 1)
        self.assertEqual(len(scan_file(p, self.matcher, {'masters': True}).units), 2)
        self.assertEqual(save_file(r, 'direct').status, '成功')

    def test_txt_newlines_bom_and_backup(self):
        p = self.root/'a.txt'
        original = codecs.BOM_UTF8 + '旧公司\r\n其他\r\n'.encode()
        p.write_bytes(original)
        out = save_file(scan_file(p, self.matcher), 'backup')
        self.assertEqual(out.status, '成功', out.detail)
        self.assertEqual(Path(out.backup).read_bytes(), original)
        self.assertEqual(p.read_bytes(), codecs.BOM_UTF8+'新企业集团\r\n其他\r\n'.encode())

    def test_utf16_big_endian_preserved(self):
        p = self.root/'a.txt'
        p.write_bytes(codecs.BOM_UTF16_BE+'旧公司\r\n'.encode('utf-16-be'))
        self.assertEqual(save_file(scan_file(p, self.matcher), 'direct').status, '成功')
        self.assertEqual(p.read_bytes(), codecs.BOM_UTF16_BE+'新企业集团\r\n'.encode('utf-16-be'))

    def test_explicit_chinese_encoding(self):
        p = self.root/'a.txt'
        p.write_bytes('旧公司'.encode('gb18030'))
        self.assertTrue(scan_file(p, self.matcher).error)
        r = scan_file(p, self.matcher, {'encoding': 'gb18030'})
        self.assertEqual(save_file(r, 'direct').status, '成功')
        self.assertEqual(p.read_bytes().decode('gb18030'), '新企业集团')

    def test_csv_quoted_newline_and_separator_replacement(self):
        p = self.root/'a.csv'
        p.write_bytes('名称,备注\r\n旧公司,"第一行\n第二行,内容"\r\n'.encode())
        m = Matcher([Rule('旧公司', '新,企业')])
        r = scan_file(p, m)
        self.assertEqual(save_file(r, 'direct').status, '成功')
        rows = list(csv.reader(io.StringIO(p.read_bytes().decode(), newline='')))
        self.assertEqual(rows[1], ['新,企业', '第一行\n第二行,内容'])

    def test_file_changed_refuses_overwrite(self):
        p = self.root/'a.txt'
        p.write_text('旧公司')
        r = scan_file(p, self.matcher)
        p.write_text('用户更新')
        out = save_file(r, 'direct')
        self.assertEqual(out.status, '失败')
        self.assertEqual(p.read_text(), '用户更新')

    def test_existing_copy_never_overwritten(self):
        p = self.root/'a.txt'
        p.write_text('旧公司')
        out = self.root/'out'
        out.mkdir()
        (out/'a.txt').write_text('存在')
        result = save_file(scan_file(p, self.matcher), 'copy', out)
        self.assertEqual(result.status, '失败')
        self.assertEqual((out/'a.txt').read_text(), '存在')

    def test_exclude_dedupe_no_recursion_symlink(self):
        p = self.root/'a.txt'
        p.write_text('旧公司')
        child = self.root/'child'
        child.mkdir()
        (child/'b.txt').write_text('旧公司')
        try:
            (self.root/'link').symlink_to(child, target_is_directory=True)
        except OSError:
            pass  # Windows 非管理员会话可能不具备符号链接权限。
        self.assertEqual(len(list(enumerate_files([self.root, p]))), 2)
        self.assertEqual(list(enumerate_files([self.root], recursive=False)), [p])
        self.assertEqual(list(enumerate_files([self.root], excluded=[child])), [p])

    def test_corrupt_doc_does_not_write(self):
        p = self.root/'a.docx'
        p.write_bytes(b'notzip')
        r = scan_file(p, self.matcher)
        self.assertTrue(r.error)
        self.assertFalse(r.units)
        self.assertEqual(p.read_bytes(), b'notzip')


if __name__ == '__main__':
    unittest.main()

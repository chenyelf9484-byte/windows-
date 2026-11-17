import codecs
import csv
import hashlib
import io
import os
import re
import shutil
import stat
import tempfile
import threading
import uuid
import zipfile
from collections import defaultdict
from pathlib import Path
from .model import Unit, ScanResult, SaveResult
from .ooxml import paragraph_units, excel_units, shared_strings, apply_patches, parse
from .legacy import KINDS, convert

MODERN = {'.docx', '.xlsx', '.pptx'}
TEXT = {'.txt', '.md', '.log', '.csv', '.tsv'}
IMAGES = {'.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff'}
SUPPORTED = MODERN | TEXT | IMAGES | set(KINDS)
SKIP = {'$recycle.bin', 'system volume information', 'windows', 'program files',
        'program files (x86)', '.git', '.docreplace-backups', '.docreplace-output'}


def is_link(path):
    if path.is_symlink():
        return True
    if os.name == 'nt':
        try:
            return bool(getattr(path.lstat(), 'st_file_attributes', 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT)
        except FileNotFoundError:
            return False
    return False


class Control:
    def __init__(self):
        self.cancel = threading.Event()
        self.running = threading.Event()
        self.running.set()

    def checkpoint(self):
        while not self.running.wait(.15):
            if self.cancel.is_set():
                return False
        return not self.cancel.is_set()


def fingerprint(path):
    digest = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1024*1024), b''):
            digest.update(chunk)
    return digest.hexdigest()


def enumerate_files(roots, recursive=True, excluded=(), extensions=None, max_mb=200,
                    min_mb=0, modified_after=0, control=None, on_error=None):
    seen = set()
    excluded = [Path(p).resolve() for p in excluded]
    def blocked(p):
        try:
            return is_link(p) or p.name.lower() in SKIP or p.name.startswith(('~$', '.docreplace-')) or any(
                p.resolve() == e or e in p.resolve().parents for e in excluded)
        except OSError as exc:
            if on_error:
                on_error(str(p), str(exc))
            return True
    for root in roots:
        root = Path(root).absolute()
        if blocked(root):
            continue
        if not root.exists():
            if on_error:
                on_error(str(root), '路径不存在')
            continue
        def walk_error(exc):
            if on_error:
                on_error(str(exc.filename), str(exc))
        paths = [root] if root.is_file() else (
            (Path(base)/name for base, dirs, names in os.walk(root, onerror=walk_error)
             if not blocked(Path(base))
             for _ in [dirs.__setitem__(slice(None), [d for d in dirs if recursive and not blocked(Path(base)/d)])]
             for name in names))
        for path in paths:
            if control and not control.checkpoint():
                return
            if blocked(path) or path.suffix.lower() not in (SUPPORTED if extensions is None else extensions):
                continue
            try:
                key = os.path.normcase(str(path.resolve()))
                if key in seen:
                    continue
                seen.add(key)
                info = path.stat()
                if min_mb*1024**2 <= info.st_size <= max_mb*1024**2 and info.st_mtime >= modified_after:
                    yield path
            except OSError as exc:
                if on_error:
                    on_error(str(path), str(exc))


def decode_text(data, fallback='auto'):
    for bom, encoding in [(codecs.BOM_UTF8, 'utf-8-sig'), (codecs.BOM_UTF16_LE, 'utf-16'),
                          (codecs.BOM_UTF16_BE, 'utf-16')]:
        if data.startswith(bom):
            return data.decode(encoding), encoding
    if fallback != 'auto':
        return data.decode(fallback), fallback
    try:
        text = data.decode('utf-8')
        if '\x00' in text:
            raise ValueError('检测到 NUL 字节；无 BOM 的 UTF-16 请手动选择编码')
        return text, 'utf-8'
    except UnicodeDecodeError as exc:
        raise ValueError('编码不明确，请选择 GB18030 或实际编码后重新扫描') from exc


def csv_info(text, ext):
    try:
        dialect = csv.Sniffer().sniff(text[:65536], delimiters=',;\t|')
    except csv.Error:
        dialect = csv.excel_tab if ext == '.tsv' else csv.excel
    return dialect


def text_units(text, ext, matcher):
    if ext in ('.csv', '.tsv'):
        rows = list(csv.reader(io.StringIO(text, newline=''), csv_info(text, ext)))
        return [Unit('csv:%d:%d' % (r, c), '第 %d 行 / 第 %d 列' % (r+1, c+1), value,
                     matcher.replace(value), len(matcher.matches(value)), payload=(r, c))
                for r, row in enumerate(rows) for c, value in enumerate(row) if matcher.matches(value)]
    units, offset = [], 0
    for n, line in enumerate(text.splitlines(keepends=True)):
        hits = matcher.matches(line)
        if hits:
            units.append(Unit('line:'+str(n), '第 %d 行' % (n+1), line, matcher.replace(line),
                              len(hits), payload=(offset, offset+len(line))))
        offset += len(line)
    return units


def check_zip(z):
    infos = z.infolist()
    if len(infos) > 30000 or sum(i.file_size for i in infos) > 512*1024**2:
        raise ValueError('文档解压内容超过限制（512MB / 30000 项）')
    if len(set(i.filename for i in infos)) != len(infos):
        raise ValueError('ZIP 包含重复部件名')
    if any(i.filename.startswith('_xmlsignatures/') for i in infos):
        raise ValueError('数字签名文档暂不修改，避免签名失效')


def office_units(path, matcher, options, ocr):
    units, warnings = [], []
    with zipfile.ZipFile(path) as z:
        check_zip(z)
        shared = shared_strings(z.read('xl/sharedStrings.xml').decode('utf-8-sig')) if 'xl/sharedStrings.xml' in z.namelist() else []
        for name in z.namelist():
            if (name.startswith('word/') or name.startswith('ppt/')) and name.endswith('.xml'):
                if name.startswith(('ppt/slideMasters/', 'ppt/slideLayouts/')) and not options.get('masters'):
                    continue
                try:
                    units.extend(paragraph_units(name, z.read(name).decode('utf-8-sig'), matcher))
                except ValueError as exc:
                    warnings.append(str(exc))
            elif re.match(r'xl/worksheets/[^/]+\.xml$', name):
                units.extend(excel_units(name, z.read(name).decode('utf-8-sig'), shared, matcher))
            if ocr and options.get('embedded') and '/media/' in name:
                if Path(name).suffix.lower() in IMAGES:
                    try:
                        units.extend(ocr.scan(name, z.read(name), matcher))
                    except Exception as exc:
                        warnings.append(name+': '+str(exc))
                else:
                    warnings.append(name+': 此图片格式暂不支持 OCR')
    return units, warnings


def scan_file(path, matcher, options=None, ocr=None):
    path = Path(path)
    options = dict(options or {})
    result = ScanResult(str(path), '', options=options)
    try:
        initial = fingerprint(path)
        ext = path.suffix.lower()
        if ext in MODERN:
            result.units, result.warnings = office_units(path, matcher, options, ocr)
        elif ext in TEXT:
            text, enc = decode_text(path.read_bytes(), options.get('encoding', 'auto'))
            result.options['actual_encoding'] = enc
            result.units = text_units(text, ext, matcher)
        elif ext in IMAGES:
            if ocr and options.get('standalone'):
                result.units = ocr.scan(path.name, path.read_bytes(), matcher)
        elif ext in KINDS:
            if ext in ('.wps', '.et', '.dps'):
                raise ValueError('WPS 专用格式原格式回写尚未验证，当前版本仅报告，不修改')
            with tempfile.TemporaryDirectory(prefix='docreplace-') as folder:
                source = Path(folder)/path.name
                shutil.copy2(path, source)
                modern = Path(folder)/('converted'+KINDS[ext][1])
                convert(source, modern, KINDS[ext][0])
                result.units, result.warnings = office_units(modern, matcher, options, ocr)
                result.options['_modern_data'] = modern.read_bytes()
            result.warnings.append('旧版文档将经办公软件往返转换；请先用副本验证排版')
        else:
            raise ValueError('暂不支持此文件格式')
        if fingerprint(path) != initial:
            raise ValueError('扫描期间文件已变化，请重新扫描')
        result.fingerprint = initial
    except Exception as exc:
        result.error = str(exc)
        result.units = []
    return result


def write_office(source, dest, units, options):
    changes, images = defaultdict(list), defaultdict(list)
    for unit in units:
        if unit.kind == 'image':
            images[unit.payload['part']].append(unit)
        else:
            part, patches = unit.payload
            changes[part].extend(patches)
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(dest, 'w') as zout:
        check_zip(zin)
        zout.comment = zin.comment
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename in changes:
                xml = apply_patches(data.decode('utf-8-sig'), changes[info.filename])
                parse(xml)
                data = xml.encode('utf-8')
            if info.filename in images:
                from .ocr import render
                data = render(data, images[info.filename], options.get('font'),
                              options.get('foreground', '#000000'), options.get('background', '#ffffff'))
            zout.writestr(info, data)
    with zipfile.ZipFile(dest) as z:
        bad = z.testzip()
        if bad:
            raise ValueError('写入后校验失败：'+bad)


def write_text(source, dest, result, units):
    data = source.read_bytes()
    text, _ = decode_text(data, result.options['actual_encoding'])
    if source.suffix.lower() in ('.csv', '.tsv'):
        dialect = csv_info(text, source.suffix.lower())
        rows = list(csv.reader(io.StringIO(text, newline=''), dialect))
        for unit in units:
            r, c = unit.payload
            rows[r][c] = unit.after
        stream = io.StringIO(newline='')
        newline = '\r\n' if '\r\n' in text else ('\r' if '\r' in text else '\n')
        csv.writer(stream, dialect, lineterminator=newline).writerows(rows)
        changed = stream.getvalue()
        if not text.endswith(('\r', '\n')):
            changed = changed[:-len(newline)]
    else:
        changed = apply_patches(text, [(u.payload[0], u.payload[1], u.after) for u in units])
    enc = result.options['actual_encoding']
    encoded = changed.encode(enc)
    if enc == 'utf-16' and data.startswith(codecs.BOM_UTF16_BE):
        encoded = codecs.BOM_UTF16_BE + changed.encode('utf-16-be')
    # 编码必须可以往返还原。
    if encoded.decode(enc) != changed:
        raise ValueError('编码校验失败')
    dest.write_bytes(encoded)


def save_file(result, mode='copy', output=None, root=None):
    source = Path(result.path)
    record = SaveResult(str(source), '失败')
    units = [u for u in result.units if u.selected]
    if not units:
        record.status = '跳过'
        record.detail = '没有勾选修改项目'
        return record
    temp_path = None
    try:
        if result.error:
            raise ValueError(result.error)
        if mode not in ('copy', 'backup', 'direct'):
            raise ValueError('保存模式无效')
        if is_link(source):
            raise ValueError('不修改符号链接')
        if fingerprint(source) != result.fingerprint:
            raise ValueError('文件在扫描后已变化，请重新扫描')
        if source.with_name('~$'+source.name).exists() or source.with_name('~$'+source.name[2:]).exists():
            raise ValueError('发现 Office 锁文件，请关闭文档后重试')
        if mode != 'copy' and not source.stat().st_mode & stat.S_IWUSR:
            raise ValueError('文件只读')
        if mode == 'copy':
            if not output:
                raise ValueError('请选择输出目录')
            base = Path(root).resolve() if root else source.parent.resolve()
            try:
                relative = source.resolve().relative_to(base)
            except ValueError:
                relative = Path(source.name)
            destination = Path(output).resolve()/relative
            if destination.resolve() == source.resolve():
                raise ValueError('另存副本不能指向原文件')
            destination.parent.mkdir(parents=True, exist_ok=True)
            if destination.exists():
                raise ValueError('输出文件已存在，请更换输出目录或文件名')
        else:
            destination = source
        fd, tmp = tempfile.mkstemp(prefix='.docreplace-', suffix=source.suffix, dir=destination.parent)
        os.close(fd)
        temp_path = Path(tmp)
        ext = source.suffix.lower()
        if ext in MODERN:
            write_office(source, temp_path, units, result.options)
        elif ext in TEXT:
            write_text(source, temp_path, result, units)
        elif ext in IMAGES:
            from .ocr import render
            temp_path.write_bytes(render(source.read_bytes(), units, result.options.get('font'),
                                         result.options.get('foreground', '#000000'), result.options.get('background', '#ffffff')))
        elif ext in KINDS:
            with tempfile.TemporaryDirectory(prefix='docreplace-') as folder:
                modern = Path(folder)/('converted'+KINDS[ext][1])
                changed = Path(folder)/('changed'+KINDS[ext][1])
                modern.write_bytes(result.options['_modern_data'])
                write_office(modern, changed, units, result.options)
                convert(changed, temp_path, KINDS[ext][0], to_modern=False)
                verify = Path(folder)/('verified'+KINDS[ext][1])
                convert(temp_path, verify, KINDS[ext][0])
                with zipfile.ZipFile(verify) as z:
                    check_zip(z)
                    if z.testzip():
                        raise ValueError('旧版格式回读校验失败')
        else:
            raise ValueError('格式不支持保存')
        with open(temp_path, 'rb+') as f:
            os.fsync(f.fileno())
        if fingerprint(source) != result.fingerprint:
            raise ValueError('保存期间原文件已变化，已取消覆盖')
        if mode == 'backup':
            backup = source.parent/'.docreplace-backups'/uuid.uuid4().hex/source.name
            backup.parent.mkdir(parents=True)
            shutil.copy2(source, backup)
            if fingerprint(backup) != result.fingerprint:
                raise ValueError('备份校验失败')
            record.backup = str(backup)
        shutil.copymode(source, temp_path)
        if mode == 'copy':
            # 排他创建避免与并行任务意外覆盖同名输出。
            try:
                with open(destination, 'xb') as out, open(temp_path, 'rb') as inp:
                    shutil.copyfileobj(inp, out)
                    out.flush()
                    os.fsync(out.fileno())
            except FileExistsError:
                raise ValueError('输出文件已存在，未覆盖')
            except Exception:
                destination.unlink(missing_ok=True)
                raise
        else:
            os.replace(temp_path, destination)
        record.status, record.destination = '成功', str(destination)
        record.detail = '已修改 %d 项 / %d 处关键词' % (len(units), sum(u.count for u in units))
    except Exception as exc:
        record.detail = str(exc)
    finally:
        if temp_path:
            temp_path.unlink(missing_ok=True)
    return record

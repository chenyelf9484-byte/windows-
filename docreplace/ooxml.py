"""对 ZIP 内 XML 的文本节点做局部替换，非目标 ZIP 项保持原始字节。

不通过整篇文档重建来写回，保留未知扩展、关系、样式和动画。
"""
import html
import re
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from .model import Unit

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
TOKEN = re.compile(r'<(?P<tag>(?:[\w.-]+:)?t)\b(?P<attrs>[^>]*?)(?<!/)>(?P<value>.*?)</(?P=tag)\s*>', re.S)
CELL = re.compile(r'<(?P<tag>(?:[\w.-]+:)?c)\b(?P<attrs>[^>]*?)(?<!/)>(?P<body>.*?)</(?P=tag)\s*>', re.S)


def parse(xml):
    if '<!DOCTYPE' in xml.upper() or '<!ENTITY' in xml.upper():
        raise ValueError('不处理包含 DTD/实体声明的 XML')
    return ET.fromstring(xml)


def xml_text(s):
    # ElementTree 只做解析验证；原始 XML 用于局部写回。
    return html.unescape(s.replace('\r\n', '\n').replace('\r', '\n'))


def apply_patches(raw, patches):
    last = len(raw)
    for start, end, value in sorted(patches, reverse=True):
        if end > last:
            raise ValueError('文档修改区域重叠')
        raw = raw[:start] + value + raw[end:]
        last = start
    return raw


def edited_nodes(texts, matches):
    """将跨 run 的命中映射回节点，保留每个未命中文字的样式。"""
    offsets, pos = [], 0
    for t in texts:
        offsets.append(pos)
        pos += len(t)
    values = list(texts)
    for start, end, target in reversed(matches):
        first = next(i for i, (p, t) in enumerate(zip(offsets, texts)) if p <= start < p + len(t))
        final = next(i for i, (p, t) in enumerate(zip(offsets, texts)) if p < end <= p + len(t))
        if first == final:
            v = values[first]
            values[first] = v[:start-offsets[first]] + target + v[end-offsets[first]:]
        else:
            values[first] = values[first][:start-offsets[first]] + target
            for i in range(first+1, final):
                values[i] = ''
            values[final] = values[final][end-offsets[final]:]
    return values


def token_patches(tokens, texts, matcher):
    values = edited_nodes(texts, matcher.matches(''.join(texts)))
    patches = []
    for tok, before, after in zip(tokens, texts, values):
        if before == after:
            continue
        attrs = tok.group('attrs')
        if 'xml:space' not in attrs:
            attrs += ' xml:space="preserve"'
        elif re.search(r'xml:space\s*=\s*[\'"]default[\'"]', attrs):
            attrs = re.sub(r'xml:space\s*=\s*[\'"]default[\'"]', 'xml:space="preserve"', attrs)
        patches.append((tok.start(), tok.end(), '<%s%s>%s</%s>' %
                        (tok.group('tag'), attrs, escape(after), tok.group('tag'))))
    return patches


def paragraph_units(part, xml, matcher):
    root = parse(xml)
    if 'purl.oclc.org/ooxml' in root.tag:
        raise ValueError('Strict OOXML 暂不支持：'+part)
    tokens = list(TOKEN.finditer(xml))
    nodes = [e for e in root.iter() if e.tag in ('{%s}t' % W, '{%s}t' % A)]
    # 自闭合空节点不产生字符，也不纳入索引。
    nodes = [e for e in nodes if e.text is not None]
    tokens = [t for t in tokens if xml_text(t.group('value')) != '']
    if len(tokens) != len(nodes):
        raise ValueError('文本节点结构暂不支持，已跳过该 XML 部件：' + part)
    indexed = {id(node): token for node, token in zip(nodes, tokens)}
    # 使用段落自身的遍历，嵌套段落独立处理；显式分隔 run 由下方分组完成。
    groups = []
    def collect(node, active):
        if node.tag in ('{%s}p' % W, '{%s}p' % A):
            active = [None]
        if node.tag in ('{%s}tab' % W, '{%s}br' % W, '{%s}br' % A, '{%s}fldChar' % W):
            active[0] = None
        if id(node) in indexed:
            if active[0] is None:
                active[0] = []
                groups.append(active[0])
            active[0].append(indexed[id(node)])
        for child in node:
            collect(child, active)
    collect(root, [None])
    out = []
    for n, group in enumerate(groups):
        texts = [xml_text(t.group('value')) for t in group]
        before = ''.join(texts)
        hits = matcher.matches(before)
        if hits:
            out.append(Unit(part+':p'+str(n), part+' / 段落 '+str(n+1), before,
                            matcher.replace(before), len(hits), payload=(part, token_patches(group, texts, matcher))))
    return out


def excel_units(part, xml, shared, matcher):
    root = parse(xml)
    if 'purl.oclc.org/ooxml' in root.tag:
        raise ValueError('Strict OOXML 暂不支持：'+part)
    namespaces = dict(re.findall(r'\b(xmlns(?::[\w.-]+)?)\s*=\s*[\'"]([^\'"]*)[\'"]', xml))
    namespace_attrs = ' '.join('%s="%s"' % (k, v) for k, v in namespaces.items())
    out = []
    for index, cell in enumerate(CELL.finditer(xml)):
        attrs, body = cell.group('attrs'), cell.group('body')
        if re.search(r'<(?:[\w.-]+:)?f(?:\s|>|/)', body):
            continue
        typ = re.search(r'\bt\s*=\s*[\'"]([^\'"]+)', attrs)
        typ = typ.group(1) if typ else ''
        if typ == 's':
            v = re.search(r'<(?:[\w.-]+:)?v\b[^>]*>(\d+)</', body)
            if not v:
                continue
            textxml = shared[int(v.group(1))]
        elif typ == 'inlineStr':
            textxml = body
        elif typ == 'str':
            v = re.search(r'<(?:[\w.-]+:)?v\b[^>]*>(.*?)</', body, re.S)
            if not v:
                continue
            textxml = '<t>%s</t>' % v.group(1)
        else:
            continue
        ts = list(TOKEN.finditer(textxml))
        texts = [xml_text(t.group('value')) for t in ts]
        before = ''.join(texts)
        hits = matcher.matches(before)
        if not hits:
            continue
        changed = apply_patches(textxml, token_patches(ts, texts, matcher))
        # 转换为当前工作表命名空间的 inlineStr，sharedStrings 本身不修改。
        children = list(parse(changed)) if typ == 's' else None
        if typ == 's':
            inner = ''.join(ET.tostring(e, encoding='unicode') for e in children)
        elif typ == 'inlineStr':
            isnode = parse('<wrapper %s>%s</wrapper>' % (namespace_attrs, changed))[0]
            inner = ''.join(ET.tostring(e, encoding='unicode') for e in isnode)
        else:
            inner = '<t xml:space="preserve">%s</t>' % escape(matcher.replace(before))
        prefix = cell.group('tag')[:-1]
        attrs = re.sub(r'\s+t\s*=\s*([\'"]).*?\1', '', attrs)
        replacement = '<%s%s t="inlineStr"><%sis>%s</%sis></%s>' % (
            cell.group('tag'), attrs, prefix, inner, prefix, cell.group('tag'))
        address = re.search(r'\br\s*=\s*[\'"]([^\'"]+)', attrs)
        label = address.group(1) if address else str(index+1)
        out.append(Unit(part+':'+label, part+' / '+label, before, matcher.replace(before), len(hits),
                        payload=(part, [(cell.start(), cell.end(), replacement)])))
    return out


def shared_strings(xml):
    root = parse(xml)
    return [ET.tostring(si, encoding='unicode') for si in root if si.tag.endswith('}si')]

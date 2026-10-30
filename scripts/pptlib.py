# -*- coding: utf-8 -*-
"""Shared helpers: namespace-aware pptx access, shape walking, math linearization."""
import os
import re
import zipfile
from lxml import etree

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
A = '{%s}' % NS['a']
P = '{%s}' % NS['p']
M = '{%s}' % NS['m']
A14 = '{%s}' % NS['a14']
MC = '{%s}' % NS['mc']

EMU_PER_INCH = 914400.0


def slide_names(zf):
    names = [n for n in zf.namelist() if re.match(r'^ppt/slides/slide\d+\.xml$', n)]
    return sorted(names, key=lambda s: int(re.findall(r'(\d+)\.xml$', s)[0]))


class Deck:
    """One .pptx file: slide XML plus the raw zip for relationship lookup."""

    def __init__(self, path):
        self.path = path
        self.name = os.path.basename(path)
        self.zip = zipfile.ZipFile(path)
        self.slides = []
        self.errors = []
        for idx, name in enumerate(slide_names(self.zip), 1):
            try:
                self.slides.append((idx, etree.fromstring(self.zip.read(name))))
            except Exception as exc:
                self.errors.append((idx, name, repr(exc)))

    def __len__(self):
        return len(self.slides)

    def close(self):
        try:
            self.zip.close()
        except Exception:
            pass


def deck_files(decks_dir):
    """All pptx in a directory, sorted, skipping PowerPoint lock files."""
    out = []
    for f in sorted(os.listdir(decks_dir)):
        if not f.lower().endswith('.pptx'):
            continue
        if f.startswith('~$'):
            continue
        out.append(f)
    return out


def iter_shapes(root):
    """Yield every p:sp / p:pic / p:graphicFrame, descending into groups and
    unwrapping mc:AlternateContent so hoisted content is still found."""
    cSld = root.find(P + 'cSld')
    if cSld is None:
        return
    tree = cSld.find(P + 'spTree')
    if tree is None:
        return
    yield from _walk(tree)


def _walk(node):
    for child in node:
        if not isinstance(child.tag, str):
            continue
        ln = etree.QName(child).localname
        if ln in ('sp', 'pic', 'graphicFrame', 'cxnSp'):
            yield child
        elif ln == 'grpSp':
            yield child
            yield from _walk(child)
        elif ln == 'AlternateContent':
            branch = child.find(MC + 'Choice')
            if branch is None:
                branch = child.find(MC + 'Fallback')
            if branch is not None:
                yield from _walk(branch)


def shape_paras(sp):
    """Paragraphs owned by THIS shape only.

    A p:sp has exactly one p:txBody child; a graphicFrame owns cells whose
    a:txBody holds the text. Descendant search would double count inside groups.
    """
    out = []
    ln = etree.QName(sp).localname
    if ln in ('sp', 'cxnSp'):
        body = sp.find(P + 'txBody')
        if body is not None:
            out.extend(body.findall(A + 'p'))
    elif ln == 'graphicFrame':
        for cell in sp.iter(A + 'tc'):
            body = cell.find(A + 'txBody')
            if body is not None:
                out.extend(body.findall(A + 'p'))
    return out


def shape_name(sp):
    cnv = sp.find('.//' + P + 'cNvPr')
    return cnv.get('name', '') if cnv is not None else ''


def shape_geo(sp):
    """(x, y, w, h) in inches, or None."""
    xf = sp.find('.//' + A + 'xfrm')
    if xf is None:
        return None
    off, ext = xf.find(A + 'off'), xf.find(A + 'ext')
    if off is None or ext is None:
        return None
    try:
        return (int(off.get('x')) / EMU_PER_INCH, int(off.get('y')) / EMU_PER_INCH,
                int(ext.get('cx')) / EMU_PER_INCH, int(ext.get('cy')) / EMU_PER_INCH)
    except (TypeError, ValueError):
        return None


def is_hidden(sp):
    cnv = sp.find('.//' + P + 'cNvPr')
    return cnv is not None and cnv.get('hidden') == '1'


def slide_size(deck, slide_index):
    """(width, height) in inches for the presentation."""
    if not hasattr(deck, '_size'):
        pres = etree.fromstring(deck.zip.read('ppt/presentation.xml'))
        sz = pres.find(P + 'sldSz')
        deck._size = (int(sz.get('cx')) / EMU_PER_INCH,
                      int(sz.get('cy')) / EMU_PER_INCH)
    return deck._size


# ---------------------------------------------------------------- math text

def _m_text(el):
    return ''.join((t.text or '') for t in el.iter(M + 't'))


def omml_linear(el, depth=0):
    """Linearize an OMML subtree into readable text for reporting.

    The output is only for humans reading a defect report. It is never written
    back into the file.
    """
    if depth > 40 or not isinstance(el.tag, str):
        return ''
    q = el.tag

    def kid(name):
        return el.find(M + name)

    def rec(node):
        return '' if node is None else omml_linear(node, depth + 1)

    if q in (M + 'oMath', M + 'oMathPara', M + 'e', M + 'num', M + 'den',
             M + 'sub', M + 'sup', M + 'lim', M + 'fName', M + 'deg', M + 'box'):
        return ''.join(rec(c) for c in el if isinstance(c.tag, str))
    if q == M + 't':
        return el.text or ''
    if q == M + 'r':
        return _m_text(el)
    if q == M + 'f':
        return '(%s)/(%s)' % (rec(kid('num')), rec(kid('den')))
    if q == M + 'sSup':
        return '%s^(%s)' % (rec(kid('e')), rec(kid('sup')))
    if q == M + 'sSub':
        return '%s_(%s)' % (rec(kid('e')), rec(kid('sub')))
    if q == M + 'sSubSup':
        return '%s_(%s)^(%s)' % (rec(kid('e')), rec(kid('sub')), rec(kid('sup')))
    if q == M + 'rad':
        return 'sqrt(%s)' % rec(kid('e'))
    if q == M + 'd':
        beg, end = '(', ')'
        pr = kid('dPr')
        if pr is not None:
            b, e2 = pr.find(M + 'begChr'), pr.find(M + 'endChr')
            if b is not None:
                beg = b.get(M + 'val', '')
            if e2 is not None:
                end = e2.get(M + 'val', '')
        return beg + '|'.join(omml_linear(c, depth + 1) for c in el.findall(M + 'e')) + end
    if q == M + 'limLow':
        return '%s_(%s)' % (rec(kid('e')), rec(kid('lim')))
    if q == M + 'limUpp':
        return '%s^(%s)' % (rec(kid('e')), rec(kid('lim')))
    if q == M + 'func':
        return '%s %s' % (rec(kid('fName')), rec(kid('e')))
    if q == M + 'nary':
        return 'NARY(%s)' % ''.join(rec(c) for c in el if isinstance(c.tag, str))
    if etree.QName(el).localname.endswith('Pr'):
        return ''
    return ''.join(omml_linear(c, depth + 1) for c in el if isinstance(c.tag, str))


def para_segments(p):
    """Split a paragraph into ordered segments.

    Each segment is {'kind': 'text'|'math', 'text': str, 'el': element}.
    Paragraph-level mc:AlternateContent is resolved to its Choice branch only,
    because the Fallback is a redundant plain-text mirror of the same content.
    """
    segs = []
    for child in p:
        if not isinstance(child.tag, str):
            continue
        ln = etree.QName(child).localname
        if ln == 'r':
            t = child.find(A + 't')
            s = (t.text or '') if t is not None else ''
            if s:
                segs.append({'kind': 'text', 'text': s, 'el': child})
        elif ln == 'br':
            segs.append({'kind': 'text', 'text': '\n', 'el': child})
        elif child.tag == A14 + 'm':
            for om in child:
                if om.tag in (M + 'oMath', M + 'oMathPara'):
                    segs.append({'kind': 'math', 'text': omml_linear(om), 'el': om})
                    break
        elif ln == 'AlternateContent':
            choice = child.find(MC + 'Choice')
            if choice is not None:
                segs.extend(_para_segments_from(choice))
    return segs


def _para_segments_from(container):
    segs = []
    for el in container.iter():
        if not isinstance(el.tag, str):
            continue
        if el.tag == M + 'oMath':
            parent = el.getparent()
            if parent is not None and parent.tag == M + 'oMathPara':
                continue
            segs.append({'kind': 'math', 'text': omml_linear(el), 'el': el})
        elif el.tag == A + 'r':
            t = el.find(A + 't')
            s = (t.text or '') if t is not None else ''
            if s:
                segs.append({'kind': 'text', 'text': s, 'el': el})
    return segs


def para_text(p):
    return ''.join(s['text'] for s in para_segments(p))


def para_plain(p):
    return ''.join(s['text'] for s in para_segments(p) if s['kind'] == 'text')


def shape_text(sp):
    return ''.join(para_text(p) for p in shape_paras(sp)).strip()

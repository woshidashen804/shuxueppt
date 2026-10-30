# -*- coding: utf-8 -*-
"""Build a single-transform variant of one deck, for controlled experiments.

One transform at a time is the whole point: if you change two things and the
rendering improves, you still do not know which one fixed it.

    mode u   only unwrap blank-Fallback AlternateContent
    mode f   only flatten the illegal oMathPara nest
    mode uf  both

Measured result on this project: 'u' changes WPS ink by exactly 0 on every deck
tested; 'f' is what restores the formulas. Keep 'u' only for readers that truly
take the fallback branch.

Usage:
    python build_variant.py <src.pptx> <dst.pptx> <u|f|uf>
"""
import os
import shutil
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
from pptlib import A, M, A14, MC

A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main'
MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'


def declare_root_namespaces(root):
    nsmap = dict(root.nsmap)
    nsmap['a14'] = A14_NS
    nsmap['mc'] = MC_NS
    new_root = etree.Element(root.tag, nsmap=nsmap)
    for k, v in root.attrib.items():
        new_root.set(k, v)
    for child in list(root):
        new_root.append(child)
    ign = new_root.get('{%s}Ignorable' % MC_NS)
    tokens = (ign or '').split()
    if 'a14' not in tokens:
        tokens.append('a14')
    new_root.set('{%s}Ignorable' % MC_NS, ' '.join(tokens))
    return new_root


def unwrap_blank_fallback(root):
    n = 0
    for ac in list(root.iter(MC + 'AlternateContent')):
        parent = ac.getparent()
        if parent is None or etree.QName(parent).localname not in ('spTree', 'grpSp'):
            continue
        choice = ac.find(MC + 'Choice')
        if choice is None or choice.get('Requires') != 'a14':
            continue
        kids = [k for k in choice if isinstance(k.tag, str)]
        if not kids or any(etree.QName(k).localname != 'sp' for k in kids):
            continue
        fb = ac.find(MC + 'Fallback')
        ft = ''.join((e.text or '') for e in fb.iter(A + 't')).strip() \
            if fb is not None else ''
        fm = ''.join((e.text or '') for e in fb.iter(M + 't')).strip() \
            if fb is not None else ''
        if ft or fm:
            continue
        idx = list(parent).index(ac)
        for j, k in enumerate(kids):
            parent.insert(idx + j, k)
        parent.remove(ac)
        n += 1
    return n


def flatten_nest(root):
    n = 0
    for outer in list(root.iter(A14 + 'm')):
        para = outer.find(M + 'oMathPara')
        if para is None:
            continue
        inner = para.find(A14 + 'm')
        if inner is None:
            continue
        om = inner.find(M + 'oMath')
        if om is None:
            continue
        for child in list(outer):
            outer.remove(child)
        outer.append(om)
        n += 1
    return n


def build(src, dst, do_unwrap, do_flatten):
    zin = zipfile.ZipFile(src)
    tmp = tempfile.mkdtemp()
    out = os.path.join(tmp, 'o.zip')
    cu = cf = 0
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('ppt/slides/slide') \
                    and item.filename.endswith('.xml'):
                root = etree.fromstring(data)
                a = unwrap_blank_fallback(root) if do_unwrap else 0
                b = flatten_nest(root) if do_flatten else 0
                if a or b:
                    root = declare_root_namespaces(root)
                    data = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)
                    cu += a
                    cf += b
            zout.writestr(item, data)
    zin.close()
    shutil.copyfile(out, dst)
    shutil.rmtree(tmp, ignore_errors=True)
    return cu, cf


if __name__ == '__main__':
    if len(sys.argv) != 4:
        raise SystemExit('用法: python build_variant.py <src> <dst> <u|f|uf>')
    src, dst, mode = sys.argv[1], sys.argv[2], sys.argv[3]
    cu, cf = build(src, dst, 'u' in mode, 'f' in mode)
    print('%-40s unwrap=%d flatten=%d' % (os.path.basename(dst), cu, cf))

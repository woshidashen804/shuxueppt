# -*- coding: utf-8 -*-
"""Fix the two structural defects that hide math content.

--nest  (default)  <m:oMathPara><a14:m><m:oMath> is invalid: oMathPara may only
                   contain oMathParaPr and oMath. Drop the inner a14:m.
                   MEASURED: this is the step that actually makes WPS draw the
                   formulas. Five decks, single-variable control, +0 change from
                   unwrap and all recovery from this.

--unwrap           Shape-level mc:AlternateContent whose mc:Fallback is an empty
                   shell. Replace the wrapper with the Choice shape.
                   MEASURED: zero rendering change in both PowerPoint and WPS —
                   WPS reads the Choice branch. Keep it only as insurance against
                   readers that really take the fallback (older Office, some
                   mobile viewers). Not the explanation for a WPS symptom.

Both are structural and never delete text. Every touched slide is checked so the
<m:oMath> count stays identical, and the build aborts if it would drop.

Usage:
    python fix_invisible.py --decks DIR [--nest|--unwrap|--both] [--write]
"""
import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
from pptlib import A, P, M, A14, MC, deck_files

A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main'
MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
RE_OMATH = re.compile(r'<m:oMath[ >]')


def omath_count(xml):
    return len(RE_OMATH.findall(xml))


def flatten_illegal_nest(root):
    """<a14:m><m:oMathPara><a14:m><m:oMath> -> <a14:m><m:oMath>"""
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


def unwrap_blank_fallback(root):
    """Replace shape-level AlternateContent with its Choice children."""
    n = 0
    for ac in list(root.iter(MC + 'AlternateContent')):
        parent = ac.getparent()
        if parent is None:
            continue
        if etree.QName(parent).localname not in ('spTree', 'grpSp'):
            continue
        choice = ac.find(MC + 'Choice')
        if choice is None or choice.get('Requires') != 'a14':
            continue
        kids = [k for k in choice if isinstance(k.tag, str)]
        if not kids or any(etree.QName(k).localname != 'sp' for k in kids):
            continue
        fallback = ac.find(MC + 'Fallback')
        ft = ''.join((e.text or '') for e in fallback.iter(A + 't')).strip() \
            if fallback is not None else ''
        fm = ''.join((e.text or '') for e in fallback.iter(M + 't')).strip() \
            if fallback is not None else ''
        if ft or fm:
            continue          # a real fallback exists, leave it alone
        idx = list(parent).index(ac)
        for j, k in enumerate(kids):
            parent.insert(idx + j, k)
        parent.remove(ac)
        n += 1
    return n


def declare_root_namespaces(root):
    """Put a14/mc on the slide root with mc:Ignorable="a14", as PowerPoint does.

    After hoisting, the a14 namespace would otherwise be re-declared on every
    moved element. Declaring once on the root plus Ignorable is the canonical form.
    """
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


def process_deck(path, write, do_nest, do_unwrap):
    zin = zipfile.ZipFile(path)
    tmp = tempfile.mkdtemp()
    out = os.path.join(tmp, 'o.zip')
    stats = {'nest': 0, 'unwrap': 0}
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('ppt/slides/slide') \
                    and item.filename.endswith('.xml'):
                xml = data.decode('utf-8')
                before = omath_count(xml)
                root = etree.fromstring(data)
                a = flatten_illegal_nest(root) if do_nest else 0
                b = unwrap_blank_fallback(root) if do_unwrap else 0
                if a or b:
                    root = declare_root_namespaces(root)
                    data = etree.tostring(root, xml_declaration=True,
                                          encoding='UTF-8', standalone=True)
                    after = omath_count(data.decode('utf-8'))
                    if after != before:
                        raise SystemExit(
                            'ABORT %s %s: oMath %d -> %d'
                            % (os.path.basename(path), item.filename, before, after))
                    stats['nest'] += a
                    stats['unwrap'] += b
            zout.writestr(item, data)
    zin.close()
    if write:
        shutil.copyfile(out, path)
    shutil.rmtree(tmp, ignore_errors=True)
    return stats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--decks', required=True)
    ap.add_argument('--nest', action='store_true',
                    help='拆 oMathPara 内的 a14:m（默认，真正修复 WPS）')
    ap.add_argument('--unwrap', action='store_true',
                    help='展开空 Fallback（对 WPS 无实测效果，仅作兼容保险）')
    ap.add_argument('--both', action='store_true')
    ap.add_argument('--write', action='store_true')
    args = ap.parse_args()

    if args.both:
        do_nest = do_unwrap = True
    elif args.nest or args.unwrap:
        do_nest, do_unwrap = args.nest, args.unwrap
    else:
        do_nest, do_unwrap = True, False        # safe default

    print('模式：nest=%s unwrap=%s' % (do_nest, do_unwrap))
    total = {'nest': 0, 'unwrap': 0}
    for f in deck_files(args.decks):
        st = process_deck(os.path.join(args.decks, f), args.write,
                          do_nest, do_unwrap)
        if st['nest'] or st['unwrap']:
            print('%-44s 嵌套=%3d 展开=%3d' % (f, st['nest'], st['unwrap']))
        for k in total:
            total[k] += st[k]
    print('合计 嵌套=%d 展开=%d  %s'
          % (total['nest'], total['unwrap'], 'WRITE' if args.write else 'DRY-RUN'))


if __name__ == '__main__':
    main()

# -*- coding: utf-8 -*-
"""Extract figure panels from a paper PDF as PNG files.

    python extract_figures.py --pdf paper.pdf --list
    python extract_figures.py --pdf paper.pdf --out figures --dpi 300 --pages 4,5
    python extract_figures.py --pdf paper.pdf --out figures --page 4 \
        --crop 0.05,0.30,0.95,0.75 --name fig2a

``--crop`` takes fractions of the page box (left,top,right,bottom) so the same
numbers work at any DPI.  Nothing is re-drawn: the panels come straight from the
source PDF, which is what keeps coordinates, legends and statistics intact.
"""
from __future__ import annotations

import argparse
from hashlib import file_digest
import json
import os
from pathlib import Path
import sys

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover
    sys.exit('PyMuPDF is required: python -m pip install pymupdf')


def page_list(text, page_count):
    if not text:
        return list(range(page_count))
    out = []
    for part in text.split(','):
        part = part.strip()
        if '-' in part:
            a, b = part.split('-')
            out.extend(range(int(a) - 1, int(b)))
        else:
            out.append(int(part) - 1)
    bad = [p + 1 for p in out if p < 0 or p >= page_count]
    if bad:
        sys.exit('pages out of range: %s (document has %d)' % (bad, page_count))
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--pdf', required=True)
    ap.add_argument('--out', default='figures')
    ap.add_argument('--dpi', type=int, default=300)
    ap.add_argument('--pages', help='1-based, e.g. 4,5,7-9 (default: all)')
    ap.add_argument('--page', type=int, help='single page for --crop')
    ap.add_argument('--crop', help='left,top,right,bottom as fractions of the page')
    ap.add_argument('--name', help='output stem for --crop')
    ap.add_argument('--list', action='store_true', help='print page boxes and exit')
    ap.add_argument('--provenance', type=Path, help='write a new local JSON record for --crop')
    ap.add_argument('--source', help='anonymous figure reference, e.g. Figure 2A')
    ap.add_argument('--panel', help='optional panel label, e.g. A')
    args = ap.parse_args()
    if args.provenance and (not args.crop or not (args.source or '').strip()):
        ap.error('--provenance requires --crop and a nonempty --source')
    if args.provenance and args.provenance.exists():
        ap.error('--provenance must be a new file')

    doc = fitz.open(args.pdf)
    if args.list:
        for i, page in enumerate(doc, 1):
            r = page.rect
            print('p%-3d %.1f x %.1f pt   %d images' % (i, r.width, r.height, len(page.get_images())))
        return

    os.makedirs(args.out, exist_ok=True)
    zoom = args.dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    if args.crop:
        if not args.page or not args.name:
            sys.exit('--crop needs --page and --name')
        if not 1 <= args.page <= len(doc):
            ap.error('--page is outside the PDF')
        page = doc[args.page - 1]
        try:
            l, t, r, b = [float(x) for x in args.crop.split(',')]
        except ValueError:
            ap.error('--crop needs four numeric page fractions')
        if not (0 <= l < r <= 1 and 0 <= t < b <= 1):
            ap.error('--crop must be an ordered rectangle within the PDF page')
        box = fitz.Rect(l * page.rect.width, t * page.rect.height,
                        r * page.rect.width, b * page.rect.height)
        pix = page.get_pixmap(matrix=mat, clip=box, alpha=False)
        path = os.path.join(args.out, args.name + '.png')
        if args.provenance and args.provenance.resolve() == Path(path).resolve():
            ap.error('--provenance must be separate from the PNG')
        pix.save(path)
        if args.provenance:
            with open(path, 'rb') as file:
                digest = file_digest(file, 'sha256').hexdigest()
            data = {'source': args.source, 'pdf_page': args.page,
                    'crop': [l, t, r, b], 'sha256': digest}
            if args.panel:
                data['panel'] = args.panel
            with args.provenance.open('x', encoding='utf-8') as file:
                json.dump(data, file, ensure_ascii=False, indent=2)
                file.write('\n')
        print('%s  %.0f x %.0f px' % (path, pix.width, pix.height))
        return

    for idx in page_list(args.pages, len(doc)):
        page = doc[idx]
        pix = page.get_pixmap(matrix=mat, alpha=False)
        path = os.path.join(args.out, 'p%02d.png' % (idx + 1))
        pix.save(path)
        print('%s  %.0f x %.0f px' % (path, pix.width, pix.height))


if __name__ == '__main__':
    main()

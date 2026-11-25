#!/usr/bin/env python3
"""Type a GRAFIX / STRAVIX Tissue Transport Log (S+N form MSEE7-34991-0722) on the blank template.

  python3 -I tools/transfer_log.py SPEC.json OUT.pdf [--jpg OUT.jpg]

SPEC.json is a list of 1-3 rows (the form has three), each:
  {"lot": "LYU-251827", "sn": "24004",          lot and unit number, one per line
   "source": ["210494", "Sherman"],             source facility: acct # (+ short name)
   "pl": true,                                  lyopreserved (PL): cryo container = N/A (PL)
   "undamaged": true, "labeling": false,        shipper integrity boxes (cryo: "cryo": true/false when not PL)
   "sent": ["11/6/26", "7am"], "by": "AJ Peterson",
   "to": ["Alexian Brothers", "Hospital", "186035"],
   "received": ["11/6/26", "8am"], "received_by": "Dr. Lucas",
   "reason": ["Needed", "same day"]}

AJ's rule: reason is always "Needed same day"; transported = the day the unit was used, 7am; received = the same
day, 8am; received by the surgeon. Save it as case-sheets/<initials>-<mmdd>-transport-log.pdf/.jpg, link it from
the case sheet page ("## Transfer log") and add the unit to the tracker's Transfers tab.
"""
import io, json, os, subprocess, sys
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, 'transport_log_template.pdf')
ROW_DY = [0, 159.6, 319.2]                      # the three rows of the form, top to bottom
COLS = {'lot': 43, 'source': 145, 'sent': 487, 'by': 582, 'to': 687, 'received': 811, 'received_by': 912, 'reason': 1078}
BOX_X = {'yes': 234, 'no': 310, 'na': 409}
BOX_Y = {'undamaged': 552, 'cryo': 499, 'labeling': 448}


def lines(v):
    return [str(x) for x in (v if isinstance(v, (list, tuple)) else [v]) if x not in (None, '')]


def build(rows, out_pdf, jpg=None):
    assert 1 <= len(rows) <= 3, 'the form has three rows'
    base = PdfReader(TEMPLATE)
    W, H = float(base.pages[0].mediabox.width), float(base.pages[0].mediabox.height)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(W, H))
    c.setFillColorRGB(.08, .08, .08)
    c.setStrokeColorRGB(.05, .05, .05)
    for r, dy in zip(rows, ROW_DY):
        def put(key, val, size=13, y0=563.5, lead=17):
            for i, t in enumerate(lines(val)):
                c.setFont('Helvetica', size)
                c.drawString(COLS[key], y0 - dy - i * lead, t)
        put('lot', [r.get('lot'), r.get('sn')])
        src = lines(r.get('source'))
        put('source', src[:1])
        if src[1:]:
            c.setFont('Helvetica', 10.5)
            for i, t in enumerate(src[1:]):
                c.drawString(COLS['source'], 548 - dy - i * 13, t)
        put('sent', r.get('sent'))
        put('by', r.get('by'))
        put('to', r.get('to'), size=11.5, y0=565, lead=15)
        put('received', r.get('received'))
        put('received_by', r.get('received_by'))
        put('reason', r.get('reason'))
        marks = [('undamaged', 'yes' if r.get('undamaged', True) else 'no'),
                 ('cryo', 'na' if r.get('pl') else ('yes' if r.get('cryo', True) else 'no')),
                 ('labeling', 'yes' if r.get('labeling') else 'no')]
        c.setLineWidth(1.6)
        for row, col in marks:
            x, y = BOX_X[col], BOX_Y[row] - dy
            c.line(x - 5.3, y - 5.3, x + 5.3, y + 5.3)
            c.line(x - 5.3, y + 5.3, x + 5.3, y - 5.3)
    c.save()
    buf.seek(0)
    w = PdfWriter()
    w.append(base)
    w.pages[0].merge_page(PdfReader(buf).pages[0])
    with open(out_pdf, 'wb') as f:
        w.write(f)
    if jpg:
        stem = jpg[:-4] if jpg.lower().endswith('.jpg') else jpg
        subprocess.run(['pdftoppm', '-r', '110', '-jpeg', '-jpegopt', 'quality=88', '-singlefile', out_pdf, stem],
                       check=True, stderr=subprocess.DEVNULL)


if __name__ == '__main__':
    a = sys.argv[1:]
    jpg = a[a.index('--jpg') + 1] if '--jpg' in a else None
    build(json.load(open(a[0])), a[1], jpg)
    print('wrote', a[1], jpg or '')

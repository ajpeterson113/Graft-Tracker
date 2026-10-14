#!/usr/bin/env python3
"""Make a filled S+N case sheet (JPG for the web page + PDF) from sticker photos.

    python3 -I tools/case_sheet.py SPEC.json OUT_BASENAME

SPEC.json:
  {"banner": "PO pending  |  Application 3",          # top strip; omit for none
   "acct": "Acct #253985",                             # optional, top centre (like AJ's handwritten sheets)
   "po": "",                                           # Purchase Order line ("" while waiting)
   "hospital": "...", "street": "...", "city": "...",
   "rep": "AJ Peterson", "surgeon": "Dr. Joshi", "patient": "M.L.",   # initials only
   "dos": "10-13-26", "price": "$360.00",
   "stickers": [{"image": "sticker.png", "crop": [x0, y0, x1, y1]}],  # crop optional
   "sig_date": "10-13-26", "email": "ajoseph.peterson@smith-nephew.com"}

Writes OUT_BASENAME.jpg (1200 px wide, for the case-sheets page) and OUT_BASENAME.pdf
(full 300 DPI). The blank form is tools/case_sheet_template.png (AJ's Case Sheet.pdf at
300 DPI). AJ's rules: use the real sticker photos (typed lot/SN only when he asks); once a PO
exists it goes on the Purchase Order line and the banner must not still say "PO pending";
multi-piece cases put the stickers large across the middle of the table, date once, and
"Each:" / "Total:" in the Unit Price column (pass "price": "Each: $x\\nTotal: $y").
"""
import json
import os
import sys

from PIL import Image, ImageDraw, ImageFont, ImageOps

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, 'case_sheet_template.png')
SANS = '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf'
BOLD = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
# Template geometry at 300 DPI (2550 x 3300): field underlines and table grid
FIELD_X = 745
FIELD_Y = {'po': 828, 'hospital': 889, 'street': 950, 'city': 1011, 'rep': 1071, 'surgeon': 1132, 'patient': 1192}
COLS = [218, 729, 1123, 1499, 1875, 2328]          # Date | Item | Lot | Ref | Unit Price
ROWS = [1434, 1609, 1795, 1985]                     # three line-item rows
PRINT_NAME = (860, 2055)
SIGN = (930, 2138)
SIGN_DATE_X = 1930
EMAIL = (640, 2220)
BANNER_H = 200


def font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except OSError:
        return ImageFont.load_default()


def text_on_line(d, x, line_y, s, f):
    """Write s so it sits just above an underline at line_y."""
    if not s:
        return
    d.text((x, line_y - 9), s, fill=(25, 25, 25), font=f, anchor='ls')   # baseline just above the line


def centred(d, box, s, f):
    lines = str(s).split('\n')
    hs = [d.textbbox((0, 0), ln, font=f) for ln in lines]
    total = sum(b[3] - b[1] for b in hs) + 18 * (len(lines) - 1)
    y = (box[1] + box[3]) / 2 - total / 2
    for ln, b in zip(lines, hs):
        w = b[2] - b[0]
        d.text(((box[0] + box[2]) / 2 - w / 2 - b[0], y - b[1]), ln, fill=(25, 25, 25), font=f)
        y += (b[3] - b[1]) + 18


def build(spec, out):
    form = Image.open(TEMPLATE).convert('RGB')
    W, H = form.size
    top = BANNER_H if spec.get('banner') else 0
    page = Image.new('RGB', (W, H + top), 'white')
    page.paste(form, (0, top))
    d = ImageDraw.Draw(page)
    if top:
        d.rectangle([0, 0, W, top - 12], fill=(252, 235, 217))
        d.rectangle([0, top - 12, W, top], fill=(233, 150, 62))
        fb = font(BOLD, 92)
        b = d.textbbox((0, 0), spec['banner'], font=fb)
        d.text((70, (top - 12) / 2 - (b[3] + b[1]) / 2), spec['banner'], fill=(80, 80, 80), font=fb)
    f = font(SANS, 54)
    if spec.get('acct'):
        d.text((900, top + 330), spec['acct'], fill=(25, 25, 25), font=font(SANS, 76))
    for k, y in FIELD_Y.items():
        text_on_line(d, FIELD_X, top + y, spec.get(k, ''), f)
    # line item: date in column 1, price in column 5, stickers across the middle
    r0, r1 = top + ROWS[0], top + ROWS[1]
    centred(d, (COLS[0], r0, COLS[1], r1), spec.get('dos', ''), f)
    centred(d, (COLS[4], r0, COLS[5], r1), spec.get('price', ''), f)
    stickers = []
    for s in spec.get('stickers') or []:
        im = ImageOps.exif_transpose(Image.open(s['image'])).convert('RGB')
        if s.get('crop'):
            im = im.crop(tuple(s['crop']))
        stickers.append(im)
    if stickers:
        n = len(stickers)
        span_w = COLS[4] - COLS[1] - 40                      # Item..Ref columns
        w = min(700, (span_w - 20 * (n - 1)) // n) if n > 1 else 700
        x = COLS[1] + 20 + (span_w - (w * n + 20 * (n - 1))) // 2
        for im in stickers:
            h = int(im.height * w / im.width)
            page.paste(im.resize((w, h), Image.LANCZOS), (x, r0 + 14))
            x += w + 20
    text_on_line(d, PRINT_NAME[0], top + PRINT_NAME[1], spec.get('rep', ''), f)
    text_on_line(d, SIGN[0], top + SIGN[1], spec.get('rep', ''), f)
    text_on_line(d, SIGN_DATE_X, top + SIGN[1], spec.get('sig_date', spec.get('dos', '')), f)
    text_on_line(d, EMAIL[0], top + EMAIL[1], spec.get('email', ''), f)
    dpi = (H + top) / 11.0                                   # fit an 11-inch page height
    page.save(out + '.pdf', 'PDF', resolution=dpi)
    small = page.resize((1200, int(page.height * 1200 / W)), Image.LANCZOS)
    small.save(out + '.jpg', quality=88)
    return out + '.jpg', out + '.pdf'


if __name__ == '__main__':
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    print('\n'.join(build(json.load(open(sys.argv[1], encoding='utf-8')), sys.argv[2])))

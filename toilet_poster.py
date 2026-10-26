"""A4 Housekeeping Request posters from an existing QR ZIP.

Uses the original QR PNGs and URLs, never rotates or recovers hashed tokens.
"""
import io
import os
import re
import zipfile
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

POSTER_REFERENCE = os.path.join(os.path.dirname(__file__), 'housekeeping_poster_master.png')
FONT_BOLD = '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
FONT_REG = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'

def _font(path, size):
    return ImageFont.truetype(path, size)

def _center(draw, xy, text, font, fill, spacing=4):
    draw.multiline_text(xy, text, font=font, fill=fill, anchor='mm', align='center', spacing=spacing)

def parse_qr_zip(data):
    """Return {toilet_id: (QR PIL image, original URL)}; require paired PNG/TXT."""
    found = {}
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        for name in z.namelist():
            if not name.lower().endswith('.txt') or '/' in name or '\\' in name:
                continue
            key = name[:-4]
            png_name = next((n for n in z.namelist() if n.lower() == (key+'.png').lower()), None)
            if not png_name:
                continue
            url = z.read(name).decode('utf-8-sig').strip()
            if '?toilet_qr=' not in url or not url.startswith(('https://', 'http://')):
                continue
            im = Image.open(io.BytesIO(z.read(png_name))).convert('RGB')
            found[key] = (im, url)
    return found

def make_page(master, qr, toilet):
    """Master is 620x891 px. QR artwork and footer text replaced."""
    im = master.convert('RGB').copy()
    d = ImageDraw.Draw(im)
    navy = '#092a48'
    # Cover only the decorative sample QR, preserving its green frame.
    d.rectangle((174, 290, 453, 553), fill='white')
    qr = qr.convert('RGB')
    # Preserve quiet zone and avoid interpolation artifacts in QR modules.
    qr.thumbnail((254, 254), Image.Resampling.NEAREST)
    im.paste(qr, (int(313-qr.width/2), int(421-qr.height/2)))
    # Replace step 2 wording with the approved multi-select language.
    d.rectangle((242, 710, 398, 760), fill='white')
    _center(d, (318, 724), '2. SELECT', _font(FONT_BOLD, 18), navy)
    _center(d, (318, 746), 'ISSUES', _font(FONT_BOLD, 18), navy)
    # Footer navy rectangle: replace the sample location lettering only.
    d.rectangle((169, 792, 580, 873), fill='#103c5b')
    tower = str(toilet.get('tower') or '').strip()
    level = str(toilet.get('level') or '').strip()
    typ = str(toilet.get('toilet_type') or toilet.get('type') or '').strip()
    name = str(toilet.get('toilet_name') or toilet.get('name') or '').strip()
    heading = f'TOWER {tower} · LEVEL {level}' if tower and level else (name or str(toilet.get('toilet_id', '')))
    subtitle = f'{typ.upper()} TOILET' if typ and 'toilet' not in typ.lower() else typ.upper()
    if not subtitle:
        subtitle = name.upper() if name.upper() != heading.upper() else str(toilet.get('toilet_id', '')).upper()
    if len(heading) > 29: heading = heading[:28] + '…'
    if len(subtitle) > 38: subtitle = subtitle[:37] + '…'
    d.text((181, 803), heading, font=_font(FONT_BOLD, 22 if len(heading)<24 else 17), fill='white')
    d.text((181, 839), subtitle, font=_font(FONT_REG, 17 if len(subtitle)<30 else 14), fill='white')
    return im

def build_pdf(qr_items, toilet_rows, selected=None):
    """Build PDF using only original supplied QR codes; skip unknown/inactive toilets."""
    master = Image.open(POSTER_REFERENCE)
    rows = {str(t['toilet_id']): t for t in toilet_rows if str(t.get('active', 'Y')).upper() not in ('N','NO','FALSE','0')}
    ids = [k for k in qr_items if k in rows and (selected is None or k in selected)]
    if not ids:
        raise ValueError('No matching active toilets with original QR images were found.')
    output = io.BytesIO()
    c = canvas.Canvas(output, pagesize=A4)
    for tid in ids:
        page = make_page(master, qr_items[tid][0], rows[tid])
        c.drawImage(ImageReader(page), 0, 0, width=A4[0], height=A4[1])
        c.showPage()
    c.save()
    return output.getvalue(), ids

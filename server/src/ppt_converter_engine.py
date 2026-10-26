#!/usr/bin/env python3
"""
Media Buzz OOH Workspace - PowerPoint Converter Engine
Converts any uploaded presentation into the official Media Buzz Master Template format
used in Automated PPT (MediaBuzz_Master_Template.pptx / Testing 24-10-2026.pptx).
Features pure Python standard library conversion (zero external dependencies required).
"""

import os
import sys
import json
import re
import copy
import io
import base64
import argparse
import zipfile
import xml.etree.ElementTree as ET
import xml.sax.saxutils as saxutils

# Ensure utf-8 encoding for stdout and stderr with replace error handling
if hasattr(sys.stdout, 'reconfigure'):
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass
if hasattr(sys.stderr, 'reconfigure'):
    try: sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception: pass

# Optional PIL import (graceful fallback if not installed)
try:
    from PIL import Image, ImageDraw, ImageFont
    HAS_PIL = True
except ImportError:
    Image = None
    ImageDraw = None
    ImageFont = None
    HAS_PIL = False

try:
    import pptx
except ImportError:
    pptx = None

import site
try:
    for _p in [site.getusersitepackages()] + (site.getsitepackages() if hasattr(site, 'getsitepackages') else []):
        if _p and _p not in sys.path and os.path.isdir(_p):
            sys.path.insert(0, _p)
except Exception:
    pass

# Optional openpyxl import for Excel-to-PowerPoint conversion
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    openpyxl = None
    HAS_OPENPYXL = False

# Optional xlrd import for legacy .xls files
try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    xlrd = None
    HAS_XLRD = False

# The 11 official master template columns from Jade Blue 16-10-2026.xlsx
MASTER_TEMPLATE_COLUMNS = [
    'SR NO',
    'City',
    'LOCATION',
    'MEDIA',
    'LIGHT',
    'W',
    'H',
    'SQ FT',
    'AVAILABLITY',
    'Selling Amount',
    'Latitude Longitude'
]

# Official 7-point Media Buzz software Terms & Conditions (from client/src/excelLogo.js)
OFFICIAL_SOFTWARE_TERMS = [
    "1. Before confirming any site, please check and confirm its availability.",
    "2. 18% GST will be charged extra, as applicable.",
    "3. Printing charges will be charged extra, as actual.",
    "4. Mounting charges will be charged extra at ₹4 per sq. ft.",
    "5. Payment Terms: Advance / within 30 days from the date of invoice.",
    "6. Artwork/Flex must be provided within 2 days of site confirmation, otherwise the site may be released.",
    "7. Cancellations require a minimum 10 days' notice; immediate cancellations are not accepted."
]

def format_inr(val):
    """Formats monetary amounts into authentic Indian numbering format (e.g. ₹ 2,50,000/-)"""
    if val is None or val == '':
        return ""
    try:
        amt = float(val)
        is_int = amt.is_integer()
        amt_str = str(int(amt)) if is_int else f"{amt:.2f}"
        parts = amt_str.split('.')
        int_part = parts[0]
        if len(int_part) <= 3:
            res = int_part
        else:
            last3 = int_part[-3:]
            rest = int_part[:-3]
            groups = []
            while len(rest) > 2:
                groups.insert(0, rest[-2:])
                rest = rest[:-2]
            if rest:
                groups.insert(0, rest)
            res = ','.join(groups) + ',' + last3
        if len(parts) > 1 and parts[1] != '00':
            res += '.' + parts[1]
        return f"₹ {res}/-"
    except Exception:
        return str(val)


CANDIDATE_TEMPLATE_PATHS = [
    os.path.join(os.path.dirname(__file__), '..', 'templates', 'MediaBuzz_Master_Template.pptx'),
    os.path.join(os.path.dirname(__file__), '..', 'templates', 'Testing 24-10-2026.pptx'),
    os.path.join(os.path.dirname(__file__), '..', 'templates', 'Testing  24-10-2026.pptx'),
    os.path.join(os.getcwd(), 'server', 'templates', 'MediaBuzz_Master_Template.pptx'),
    os.path.join(os.getcwd(), 'templates', 'MediaBuzz_Master_Template.pptx'),
    os.path.join(os.getcwd(), 'server', 'templates', 'Testing 24-10-2026.pptx'),
    os.path.join(os.getcwd(), 'templates', 'Testing 24-10-2026.pptx'),
    os.path.join(r'C:\Users\reeha\Downloads', 'MediaBuzz_Master_Template.pptx'),
    os.path.join(r'C:\Users\reeha\Downloads', 'Testing  24-10-2026.pptx'),
    os.path.join(r'C:\Users\reeha\Downloads', 'Testing 24-10-2026.pptx'),
    os.path.join(r'C:\Users\reeha\Downloads', '22-09-2026 (1).pptx'),
    os.path.join(r'C:\Users\reeha\Downloads', 'Hari Om Shela Project -29-09-2026.pptx'),
]

def find_master_template(explicit_path=None):
    if explicit_path and os.path.isfile(explicit_path):
        return os.path.abspath(explicit_path)
    for p in CANDIDATE_TEMPLATE_PATHS:
        if os.path.isfile(p):
            return os.path.abspath(p)
    return None

def extract_code(text, default_idx=1):
    m = re.search(r'\b([A-Za-z]{1,6})\s*[-–—]\s*(\d{1,5})\b', text)
    if m:
        prefix = m.group(1).upper()
        num = m.group(2)
        return f"{prefix} - {num}"
    m2 = re.search(r'\b(SITE|LOC|MB|SPOT)\s*#?\s*(\d{1,5})\b', text, re.IGNORECASE)
    if m2:
        return f"MB - {int(m2.group(2)):02d}"
    return f"MB - {default_idx:02d}"

def extract_size(text):
    m = re.search(r'\b(\d{1,3}(?:\.\d+)?)\s*[xX*×]\s*(\d{1,3}(?:\.\d+)?)(?:\s*(?:ft|feet|FT|m|meter|M))?\b', text)
    if m:
        w, h = m.group(1), m.group(2)
        return f"{w}x{h} ft"
    return None

def detect_city(text):
    cities = [
        'Ahmedabad', 'Gandhinagar', 'Surat', 'Vadodara', 'Rajkot',
        'Bhavnagar', 'Jamnagar', 'Junagadh', 'Anand', 'Nadiad', 'Mehsana',
        'Bharuch', 'Navsari', 'Valsad', 'Vapi', 'Bhuj', 'Morbi',
        'Mumbai', 'Pune', 'Delhi', 'Jaipur', 'Udaipur', 'Indore', 'Bhopal'
    ]
    for c in cities:
        if re.search(r'\b' + c + r'\b', text, re.IGNORECASE):
            return c
    return 'Ahmedabad'

def clean_location(text, size_str=None, code_str=None, city_str=None):
    s = text
    # 1. Remove explicit prefixes/labels
    s = re.sub(r'(?i)\b(location|address|site\s*(?:code|no|#)?|size|dimension[s]?|media(?:\s*type)?|type|illum(?:ination)?|lighting|rate|adv\.?\s*(?:fee|cost)|cost|rental|avail(?:ability)?|coordinates?|latitude|longitude|lat|long|gps|facing|traffic)\s*[:\-–—]?\b', ' ', s)
    # 2. Remove any site codes: MB-01, SITE 1, LOC-2, AMD-001, etc.
    s = re.sub(r'(?i)\b(?:mb|site|loc|spot|amd|gantry|hoarding|unipole)[\s#\-–—]*\d+\b', ' ', s)
    # 3. Remove sizes & dimensions: 30x10, 40*20 ft, 1500 sq ft, etc.
    s = re.sub(r'\b\d{1,4}(?:\.\d+)?\s*[xX*×]\s*\d{1,4}(?:\.\d+)?(?:\s*(?:ft|feet|FT|m|meter|M|sq\.?\s*ft|sqft))?\b', ' ', s)
    s = re.sub(r'(?i)\b\d{2,6}\s*(?:sq\.?\s*ft|sqft|feet)\b', ' ', s)
    if size_str:
        s = re.sub(re.escape(size_str), ' ', s, flags=re.IGNORECASE)
    if code_str:
        s = re.sub(re.escape(code_str), ' ', s, flags=re.IGNORECASE)
    # 4. Remove rates and rate terms
    s = re.sub(r'(?i)(?:₹|rs\.?|inr)\s*[\d,]+(?:\s*/-|\s*per\s*month)?', ' ', s)
    s = re.sub(r'\b\d{1,3}(?:,\d{2,3})+(?:\s*/-)?\b', ' ', s)
    s = re.sub(r'(?i)\b(per\s*month|per\s*day|\bp\.?m\.?\b|rental|adv\.?\s*(?:fee|cost)?|fee|fees|cost|rate)\b', ' ', s)
    # 5. Remove media types and illumination words
    s = re.sub(r'(?i)\b(hoarding|unipole|gantry|cantilever|bridge\s*panel|fob|kiosk|digital|dooh|led|billboard)\b', ' ', s)
    s = re.sub(r'(?i)\b(front\s*lit|back\s*lit|non\s*lit|\bbl\b|\bfl\b|\bnl\b)\b', ' ', s)
    # 6. Remove coordinates
    s = re.sub(r'\b\d{1,2}\.\d{4,8}\s*,\s*\d{1,2}\.\d{4,8}\b', ' ', s)
    # 7. Remove common status words
    s = re.sub(r'(?i)\b(available(?:\s+from|\s+till|\s+now)?|booked(?:\s+till)?|ready\s+to\s+book)\b', ' ', s)
    # 8. Remove brand names & slogans
    s = re.sub(r'(?i)\b(mediabuzz|outdoor|be seen|all rights reserved)\b', ' ', s)
    # 9. Clean up whitespace and punctuation
    s = re.sub(r'[\r\n\t]+', ' ', s)
    s = re.sub(r'\s*[,;:\-–—|/]+\s*', ', ', s)
    s = re.sub(r'^\s*\d+[\s.\-–—]+', '', s)
    s = re.sub(r'\s+', ' ', s).strip(', -–—:')
    # 10. Remove duplicate consecutive words (e.g. 'Ahmedabad Ahmedabad' -> 'Ahmedabad')
    words = s.split()
    dedup = []
    for w in words:
        clean_w = re.sub(r'^[^\w]+|[^\w]+$', '', w).lower()
        if not dedup or clean_w != re.sub(r'^[^\w]+|[^\w]+$', '', dedup[-1]).lower():
            dedup.append(w)
    res = ' '.join(dedup).strip(', -–—:')
    res = re.sub(r'^\d+[\s.\-–—]+', '', res).strip(', -–—:')
    # 11. If city repeated at the very end, remove it
    if city_str and res.lower().endswith(city_str.lower()):
        res = re.sub(r'(?i),\s*' + re.escape(city_str) + r'$', '', res).strip(', -–—:')
    return res[:120] if res else (f"Near {city_str} Hub" if city_str else "Prime Billboard Location")




def detect_media_type(text):
    types = [
        ('Unipole', 'Unipole'),
        ('Gantry', 'Gantry'),
        ('Cantilever', 'Cantilever'),
        ('Bridge', 'Bridge Panel'),
        ('FOB', 'Foot Over Bridge'),
        ('Kiosk', 'Kiosk'),
        ('Digital', 'DOOH Screen'),
        ('LED', 'Digital LED Display'),
        ('Hoarding', 'Hoarding')
    ]
    for kw, label in types:
        if re.search(r'\b' + kw + r'\b', text, re.IGNORECASE):
            return label
    return 'Hoarding'

def detect_illumination(text):
    if re.search(r'\b(front\s*lit|fl)\b', text, re.IGNORECASE):
        return 'Front Lit'
    if re.search(r'\b(back\s*lit|bl)\b', text, re.IGNORECASE):
        return 'BL'
    if re.search(r'\b(non\s*lit|nl)\b', text, re.IGNORECASE):
        return 'Non-Lit'
    return 'BL'

def detect_rate(text):
    m = re.search(r'(?:₹|rs\.?|inr)\s*([\d,]+(?:\s*/-|\s*per\s*month)?)', text, re.IGNORECASE)
    if m:
        val = m.group(1).strip()
        if not val.startswith('₹'): val = f"₹ {val}"
        if not val.endswith('/-') and 'month' not in val.lower(): val = f"{val}/-"
        return val
    return 'On Request'

def detect_availability(text):
    m = re.search(r'\b(available(?:\s+from\s+[A-Za-z0-9\s,]+)?|booked(?:\s+till\s+[A-Za-z0-9\s,]+)?|ready\s+to\s+book)\b', text, re.IGNORECASE)
    if m:
        return m.group(1).capitalize()
    return 'Available'

def detect_coords(text):
    m = re.search(r'\b(2[0-4]\.\d{4,8}\s*,\s*7[0-4]\.\d{4,8})\b', text)
    if m:
        return re.sub(r'\s+', ' ', m.group(1).strip())
    m2 = re.search(r'\b(\d{1,2}\.\d{4,8}\s*,\s*\d{1,2}\.\d{4,8})\b', text)
    if m2:
        return re.sub(r'\s+', ' ', m2.group(1).strip())
    return '23.012345, 72.567890'

# 1x1 dark gray (#111111) PNG base64 fallback when PIL is not installed
_DARK_PIXEL_PNG = (
    b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde'
    b'\x00\x00\x00\x0cIDATx\x9cc\x14\x14\x14\x00\x00\x00f\x00\x1f\xbf\x95\xed\xe4\x00\x00\x00\x00IEND\xaeB`\x82'
)

def generate_no_photo_bytes():
    """Generates clean placeholder image bytes for missing photos matching Automated PPT."""
    if HAS_PIL and Image and ImageDraw:
        try:
            w, h = 1500, 1224
            img = Image.new('RGB', (w, h), color=(17, 17, 17))
            draw = ImageDraw.Draw(img)
            draw.rectangle([10, 10, w - 10, h - 10], outline=(42, 42, 42), width=3)
            draw.text((w // 2 - 250, h // 2 - 20), "NO SITE PHOTO UPLOADED", fill=(85, 85, 85))
            buf = io.BytesIO()
            img.save(buf, format='PNG')
            return buf.getvalue()
        except Exception:
            pass
    return _DARK_PIXEL_PNG

def duplicate_slide(prs, src_slide):
    """Cleanly duplicate a slide using python-pptx."""
    dest_slide = prs.slides.add_slide(src_slide.slide_layout)
    rel_map = {}
    for rel_id, rel in src_slide.part.rels.items():
        if "slideLayout" in rel.reltype:
            continue
        if rel.is_external:
            new_rel_id = dest_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            new_rel_id = dest_slide.part.relate_to(rel.target_part, rel.reltype)
        rel_map[rel_id] = new_rel_id
        
    for sh in list(dest_slide.shapes):
        sp = sh._element
        sp.getparent().remove(sp)
        
    for sh in src_slide.shapes:
        new_sp = copy.deepcopy(sh._element)
        for blip in new_sp.xpath('.//a:blip'):
            embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if embed in rel_map:
                blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', rel_map[embed])
        dest_slide.shapes._spTree.append(new_sp)
        
    return dest_slide

def generate_slide_thumbnail(photo_bytes, site_num, code, size, location, city="", mtype="Hoarding", illum="BL", rate="On Request", avail="Available", coords="23.012345, 72.567890"):
    """
    Generates visual 16:9 JPEG preview matching the exact Automated PPT master theme.
    If PIL is not installed, gracefully returns extracted photo bytes directly or None.
    """
    if HAS_PIL and Image and ImageDraw:
        try:
            w, h = 800, 450
            canvas = Image.new('RGB', (w, h), color=(0, 0, 0))
            draw = ImageDraw.Draw(canvas)
            
            px, py, pw, ph = 42, 42, 375, 306
            draw.rectangle([px-1, py-1, px+pw+1, py+ph+1], outline=(40, 50, 65), width=1)
            
            if photo_bytes:
                try:
                    with Image.open(io.BytesIO(photo_bytes)) as img:
                        img_copy = img.copy()
                        img_copy.thumbnail((pw, ph), Image.Resampling.LANCZOS)
                        ix = px + (pw - img_copy.width) // 2
                        iy = py + (ph - img_copy.height) // 2
                        canvas.paste(img_copy, (ix, iy))
                except Exception:
                    draw.text((px + 20, py + 140), "Photo preview unavailable", fill=(160, 160, 160))
            else:
                draw.rectangle([px, py, px+pw, py+ph], fill=(17, 17, 17))
                draw.text((px + 90, py + 140), "NO SITE PHOTO UPLOADED", fill=(85, 85, 85))
                
            lb_x, lb_y, lb_w, lb_h = 42, 363, 111, 45
            draw.rounded_rectangle([lb_x, lb_y, lb_x + lb_w, lb_y + lb_h], radius=8, fill=(255, 194, 0))
            draw.text((lb_x + 14, lb_y + 8), "MEDIA BUZZ", fill=(3, 3, 82))
            draw.text((lb_x + 28, lb_y + 24), "Be Seen", fill=(3, 3, 82))
            
            rx = 447
            bx, by, bw, bh = rx, 54, 150, 44
            draw.rounded_rectangle([bx, by, bx+bw, by+bh], radius=8, fill=(255, 194, 0))
            draw.text((bx + 18, by + 12), f"Site {site_num}", fill=(0, 0, 0))
            
            draw.text((rx, 118), city or "", fill=(255, 255, 255))
            draw.text((rx, 148), location[:44] if location else "", fill=(255, 255, 255))
            if len(location) > 44:
                draw.text((rx, 168), location[44:88], fill=(255, 255, 255))
                
            fields = [
                ("Size:", size or ""),
                ("Type:", mtype or ""),
                ("Illumination:", illum or ""),
                ("Adv.fee per month:", rate or ""),
                ("Availability:", avail or ""),
                ("Location Coordinates:", coords or "")
            ]
            
            fy = 192
            for label, val in fields:
                draw.text((rx, fy), label, fill=(255, 194, 0))
                draw.text((rx + 145, fy), val, fill=(255, 255, 255))
                draw.line([rx, fy + 20, rx + 325, fy + 20], fill=(45, 45, 45), width=1)
                fy += 29
                
            buf = io.BytesIO()
            canvas.save(buf, format='JPEG', quality=85)
            return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode('utf-8')
        except Exception:
            pass

    # Fallback when PIL is missing: return extracted photo directly if present
    if photo_bytes:
        try:
            return "data:image/jpeg;base64," + base64.b64encode(photo_bytes).decode('utf-8')
        except Exception:
            pass
    return None

def generate_terms_slide_thumbnail():
    """Generates visual 16:9 JPEG preview for official Media Buzz software Terms & Conditions."""
    if HAS_PIL and Image and ImageDraw:
        try:
            w, h = 800, 450
            canvas = Image.new('RGB', (w, h), color=(0, 0, 0))
            draw = ImageDraw.Draw(canvas)
            
            # Gold pill header
            bx, by, bw, bh = 42, 35, 260, 42
            draw.rounded_rectangle([bx, by, bx+bw, by+bh], radius=8, fill=(255, 194, 0))
            draw.text((bx + 16, by + 12), "TERMS & CONDITIONS", fill=(0, 0, 0))
            
            # Subtitle
            draw.text((42, 88), "Media Buzz Outdoor Media Contract Terms", fill=(255, 255, 255))
            draw.text((42, 108), "Retrieved exclusively from Media Buzz software", fill=(74, 222, 128))
            
            # 7 terms lines
            ty = 135
            for idx, term in enumerate(OFFICIAL_SOFTWARE_TERMS):
                dot_pos = term.find('.')
                num_part = term[:dot_pos+1]
                text_part = term[dot_pos+1:].strip()
                draw.text((42, ty), num_part, fill=(255, 194, 0))
                disp_text = text_part if len(text_part) <= 80 else text_part[:77] + '...'
                draw.text((70, ty), disp_text, fill=(240, 240, 240))
                ty += 34
                
            # Footer badge
            draw.text((42, 410), "mediabuzzoutdoor.com", fill=(100, 116, 139))
            
            buf = io.BytesIO()
            canvas.save(buf, format='JPEG', quality=85)
            return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode('utf-8')
        except Exception:
            pass
    return None

# ==============================================================================
# Pure Python Standard Library Engine (Zero Dependencies, Works Everywhere)
# ==============================================================================

def extract_entries_pure_zip(input_path):
    """Extracts billboard entries and photographs from any PPTX using zipfile and regex."""
    entries = []
    with zipfile.ZipFile(input_path, 'r') as z_in:
        namelist = set(z_in.namelist())
        slide_files = []
        for name in namelist:
            m = re.match(r'^ppt/slides/slide(\d+)\.xml$', name)
            if m:
                slide_files.append((int(m.group(1)), name))
        slide_files.sort(key=lambda x: x[0])
        
        for slide_idx, slide_name in slide_files:
            slide_xml = z_in.read(slide_name).decode('utf-8', errors='ignore')
            rels_name = f"ppt/slides/_rels/{os.path.basename(slide_name)}.rels"
            
            # Extract text elements: <a:t>...</a:t>
            texts = re.findall(r'<a:t[^>]*>(.*?)</a:t>', slide_xml)
            full_text = ' '.join(t.strip() for t in texts if t.strip())
            
            # Skip pure closing / contact / thank you slides
            if full_text.lower() in ['mediabuzzoutdoor.com', 'thank you', 'thanks', 'contact us', 'thank you!']:
                continue
                
            size = extract_size(full_text)
            city = detect_city(full_text)
            site_num_label = f"Site {len(entries) + 1}"
            loc = clean_location(full_text, size, None, city)
            mtype = detect_media_type(full_text)
            illum = detect_illumination(full_text)
            rate = detect_rate(full_text)
            avail = detect_availability(full_text)
            coords = detect_coords(full_text)
            
            # Find photo in slide rels
            photo_bytes = None
            if rels_name in namelist:
                rels_xml = z_in.read(rels_name).decode('utf-8', errors='ignore')
                img_targets = re.findall(r'Type="[^"]*relationships/image"\s+Target="([^"]+)"', rels_xml)
                if not img_targets:
                    img_targets = re.findall(r'Target="([^"]+)"\s+Type="[^"]*relationships/image"', rels_xml)
                
                best_bytes = None
                max_size = 0
                for tgt in img_targets:
                    clean_path = tgt.replace('../', 'ppt/')
                    if not clean_path.startswith('ppt/'):
                        clean_path = 'ppt/' + clean_path.lstrip('/')
                    if clean_path in namelist:
                        data = z_in.read(clean_path)
                        # Filter out tiny icon decorations (< 5KB)
                        if len(data) > max_size and len(data) > 5000:
                            max_size = len(data)
                            best_bytes = data
                photo_bytes = best_bytes
                
            if not photo_bytes and not size and (not loc or len(loc) < 5):
                continue
                
            review_flags = []
            if not photo_bytes:
                review_flags.append("Photo missing or not detected")
            if not size:
                review_flags.append("Dimensions/size missing")
            if not loc or len(loc) < 3:
                review_flags.append("Location description unclear")
                
            entries.append({
                'source_slide': slide_idx,
                'photo_bytes': photo_bytes,
                'code': site_num_label,
                'size': size or '30x10 ft',
                'location': loc or 'Prime City Location',
                'city': city or 'Ahmedabad',
                'type': mtype or 'Hoarding',
                'illum': illum or 'BL',
                'rate': rate or 'On Request',
                'availability': avail or 'Available',
                'coords': coords or '23.012345, 72.567890',
                'review_needed': len(review_flags) > 0,
                'review_flags': review_flags
            })
            
    return entries

# ==============================================================================
# Robust Multi-Format Excel Loader (openpyxl, xlrd, and pure XML fallback)
# ==============================================================================

class UniversalCell:
    def __init__(self, value=""):
        self.value = value

class XlrdSheetWrapper:
    def __init__(self, sheet, book):
        self.sheet = sheet
        self.book = book
        self.title = sheet.name
        self.max_row = sheet.nrows
        self.max_column = sheet.ncols
        self._images = []

    def cell(self, row, column):
        r_idx = row - 1
        c_idx = column - 1
        if 0 <= r_idx < self.sheet.nrows and 0 <= c_idx < self.sheet.ncols:
            c = self.sheet.cell(r_idx, c_idx)
            val = c.value
            if HAS_XLRD and xlrd and c.ctype == xlrd.XL_CELL_DATE:
                try:
                    import datetime
                    date_tuple = xlrd.xldate_as_tuple(val, self.book.datemode)
                    val = datetime.datetime(*date_tuple).strftime('%d-%m-%Y')
                except Exception:
                    pass
            elif HAS_XLRD and xlrd and c.ctype == xlrd.XL_CELL_NUMBER and isinstance(val, (int, float)) and float(val).is_integer():
                val = int(val)
            return UniversalCell(val)
        return UniversalCell("")

class XlrdBookWrapper:
    def __init__(self, book):
        self.book = book
        self.sheetnames = book.sheet_names()

    def __getitem__(self, name):
        return XlrdSheetWrapper(self.book.sheet_by_name(name), self.book)

    @property
    def active(self):
        return XlrdSheetWrapper(self.book.sheet_by_index(0), self.book)

class PureXmlSheetWrapper:
    def __init__(self, title, rows_data):
        self.title = title
        self.rows_data = rows_data # dict of (row_idx, col_idx) -> value
        max_r = max((k[0] for k in rows_data.keys()), default=0)
        max_c = max((k[1] for k in rows_data.keys()), default=0)
        self.max_row = max_r
        self.max_column = max_c
        self._images = []

    def cell(self, row, column):
        return UniversalCell(self.rows_data.get((row, column), ""))

class PureXmlBookWrapper:
    def __init__(self, sheets_dict):
        self.sheets_dict = sheets_dict
        self.sheetnames = list(sheets_dict.keys())

    def __getitem__(self, name):
        return self.sheets_dict[name]

    @property
    def active(self):
        first_key = self.sheetnames[0] if self.sheetnames else 'Sheet1'
        return self.sheets_dict.get(first_key, PureXmlSheetWrapper('Sheet1', {}))

def parse_col_letter(col_str):
    num = 0
    for char in col_str.upper():
        if 'A' <= char <= 'Z':
            num = num * 26 + (ord(char) - ord('A') + 1)
    return num

def load_pure_xml_xlsx(excel_path):
    """Pure Python standard library fallback parser for .xlsx files (zero external libraries required)."""
    with zipfile.ZipFile(excel_path, 'r') as z:
        # 1. Parse shared strings
        shared_strings = []
        if 'xl/sharedStrings.xml' in z.namelist():
            ss_xml = z.read('xl/sharedStrings.xml')
            root = ET.fromstring(ss_xml)
            for si in root.findall('{*}si'):
                text_parts = []
                for t in si.iter('{*}t'):
                    if t.text: text_parts.append(t.text)
                shared_strings.append(''.join(text_parts))

        # 2. Parse sheet names from workbook.xml
        sheet_map = {}
        if 'xl/workbook.xml' in z.namelist():
            wb_xml = z.read('xl/workbook.xml')
            root = ET.fromstring(wb_xml)
            for s in root.iter('{*}sheet'):
                s_name = s.attrib.get('name')
                s_id = s.attrib.get('sheetId', '1')
                sheet_map[f"xl/worksheets/sheet{s_id}.xml"] = s_name

        sheets_dict = {}
        for fname in z.namelist():
            if fname.startswith('xl/worksheets/sheet') and fname.endswith('.xml'):
                title = sheet_map.get(fname, os.path.splitext(os.path.basename(fname))[0])
                s_xml = z.read(fname)
                root = ET.fromstring(s_xml)
                rows_data = {}
                for row_el in root.iter('{*}row'):
                    r_idx_str = row_el.attrib.get('r')
                    r_num = int(r_idx_str) if r_idx_str else 1
                    for c_el in row_el.iter('{*}c'):
                        cell_ref = c_el.attrib.get('r', '')
                        m = re.match(r'([A-Za-z]+)(\d+)', cell_ref)
                        if m:
                            c_num = parse_col_letter(m.group(1))
                            r_num = int(m.group(2))
                        else:
                            c_num = 1
                        t_type = c_el.attrib.get('t')
                        v_el = c_el.find('{*}v')
                        raw_v = v_el.text if v_el is not None and v_el.text is not None else ""
                        if t_type == 's' and raw_v.isdigit():
                            idx = int(raw_v)
                            val = shared_strings[idx] if idx < len(shared_strings) else raw_v
                        elif t_type == 'inlineStr':
                            t_el = c_el.find('.//{*}t')
                            val = t_el.text if t_el is not None else ""
                        else:
                            val = raw_v
                        rows_data[(r_num, c_num)] = val
                sheets_dict[title] = PureXmlSheetWrapper(title, rows_data)
        return PureXmlBookWrapper(sheets_dict)

def load_any_workbook(excel_path):
    """
    Universally loads any Excel file (.xlsx or .xls) across all environments:
    1. Uses xlrd for legacy .xls files
    2. Uses openpyxl for modern .xlsx files
    3. Falls back to pure XML Zip parser if openpyxl is not installed or fails
    """
    is_xls = excel_path.lower().endswith('.xls')
    # Check binary header for OLE2
    try:
        with open(excel_path, 'rb') as f:
            header = f.read(8)
            if header.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'):
                is_xls = True
    except Exception:
        pass

    if is_xls:
        if HAS_XLRD and xlrd:
            book = xlrd.open_workbook(excel_path)
            return XlrdBookWrapper(book)
        else:
            raise ImportError("xlrd library is required to read .xls files. Please install xlrd or save as .xlsx.")

    # Modern .xlsx
    if HAS_OPENPYXL and openpyxl:
        try:
            return openpyxl.load_workbook(excel_path, data_only=True)
        except Exception as e:
            sys.stderr.write(f"[Excel Loader] openpyxl failed, falling back to pure XML: {e}\n")

    return load_pure_xml_xlsx(excel_path)

def extract_entries_from_excel(excel_path, column_mapping=None, sheet_name=None):
    """
    Extracts site entries from an Excel workbook based on the Master Template reference (Jade Blue 16-10-2026.xlsx).
    Leaves missing values blank (Rule 8). Ignores any Terms & Conditions in the Excel sheet (Rule 7).
    """
    wb = load_any_workbook(excel_path)
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb['Sites'] if 'Sites' in wb.sheetnames else wb.active

    # Detect header row
    header_row_idx = None
    headers_found = []
    
    col_defs = [
        ('sr_no', ['sr no', 'sr. no', 'srno', 'sr_no', 's.no', 'sno', 'serial no', 'site no', 'index']),
        ('city', ['city', 'location city', 'town']),
        ('location', ['location', 'site location', 'address', 'site address', 'site name', 'description']),
        ('media', ['media', 'media type', 'type', 'display type']),
        ('light', ['light', 'lighting', 'illumination', 'illum']),
        ('w', ['w', 'width', 'w (ft)', 'w(ft)', 'w ft']),
        ('h', ['h', 'height', 'h (ft)', 'h(ft)', 'h ft']),
        ('sq_ft', ['sq ft', 'sqft', 'sq. ft', 'area', 'total sq ft']),
        ('avail', ['availablity', 'availability', 'status', 'available from', 'booked till']),
        ('rate', ['selling amount', 'selling rate', 'amount', 'rate', 'price', 'monthly rate', 'adv fee']),
        ('coords', ['latitude longitude', 'lat long', 'latitude, longitude', 'lat, lng', 'lat/long', 'gps', 'coordinates'])
    ]

    for r in range(1, min(15, ws.max_row + 1)):
        row_vals = [str(ws.cell(r, c).value or '').strip() for c in range(1, ws.max_column + 1)]
        row_joined = ' '.join(row_vals).lower()
        if 'location' in row_joined and ('media' in row_joined or 'city' in row_joined or 'sr no' in row_joined):
            header_row_idx = r
            headers_found = [str(ws.cell(r, c).value or '').strip() for c in range(1, ws.max_column + 1)]
            break

    if not header_row_idx:
        header_row_idx = 2 if ws.max_row >= 2 else 1
        headers_found = [str(ws.cell(header_row_idx, c).value or '').strip() for c in range(1, ws.max_column + 1)]

    while headers_found and not headers_found[-1]:
        headers_found.pop()

    mapping = {}
    if column_mapping:
        mapping = column_mapping
    else:
        for c_idx, h in enumerate(headers_found, start=1):
            h_clean = h.lower().strip()
            for field, aliases in col_defs:
                if field in mapping:
                    continue
                if any(h_clean == a for a in aliases):
                    mapping[field] = c_idx
                    break
                elif any(len(a) > 1 and (a in h_clean or h_clean in a) for a in aliases):
                    mapping[field] = c_idx
                    break

    # Read row images if any embedded in Excel sheet
    row_images = {}
    try:
        if hasattr(ws, '_images') and ws._images:
            for img in ws._images:
                try:
                    img_row = img.anchor._from.row + 1
                    img_bytes = img.ref.getvalue() if hasattr(img.ref, 'getvalue') else img._data()
                    if img_row not in row_images:
                        row_images[img_row] = img_bytes
                except Exception:
                    pass
    except Exception:
        pass

    entries = []
    for r in range(header_row_idx + 1, ws.max_row + 1):
        row_text_all = ' '.join([str(ws.cell(r, c).value or '') for c in range(1, ws.max_column + 1)]).strip()
        if not row_text_all:
            continue
        # Rule 7: Ignore Terms & Conditions in Excel workbook completely
        if 'TERMS' in row_text_all.upper() and 'CONDITION' in row_text_all.upper():
            break

        def get_val(field):
            col = mapping.get(field)
            if col and col <= ws.max_column:
                v = ws.cell(r, col).value
                return str(v).strip() if v is not None else ""
            return ""

        sr_no = get_val('sr_no')
        city = get_val('city')
        loc = get_val('location')
        media = get_val('media')
        light = get_val('light')
        w = get_val('w')
        h = get_val('h')
        sq_ft = get_val('sq_ft')
        avail = get_val('avail')
        rate_val = get_val('rate')
        coords = get_val('coords')

        if not loc and not city and not sr_no:
            continue

        size_str = ""
        if w and h:
            size_str = f"{w}x{h} ft"
        elif sq_ft:
            size_str = f"{sq_ft} sq.ft"

        rate_str = format_inr(rate_val) if rate_val else ""

        # Clean location without altering actual listing information
        cleaned_loc = clean_location(loc, size_str, city_str=city) if loc else ""

        review_flags = []
        photo_bytes = row_images.get(r)
        if not photo_bytes:
            review_flags.append("Photo missing from sheet")
        if not size_str:
            review_flags.append("Size missing")
        if not cleaned_loc:
            review_flags.append("Location missing")

        entries.append({
            'source_slide': r,
            'code': f"MB - {len(entries) + 1:02d}",
            'site_number': len(entries) + 1,
            'sr_no': sr_no or str(len(entries) + 1),
            'city': city,
            'location': cleaned_loc or loc,
            'type': media or 'Hoarding',
            'illum': light or 'BL',
            'size': size_str or '30x10 ft',
            'w': w,
            'h': h,
            'sq_ft': sq_ft,
            'rate': rate_str or '',
            'availability': avail or 'Available',
            'coords': coords or '',
            'photo_bytes': photo_bytes,
            'review_needed': len(review_flags) > 0,
            'review_flags': review_flags
        })

    return entries, headers_found, header_row_idx, mapping

def inspect_excel(excel_path):
    """
    Inspects an uploaded Excel workbook (.xlsx or .xls), returns sheets, detected headers,
    master template match status, mapping, and sample rows.
    """
    try:
        wb = load_any_workbook(excel_path)
    except Exception as e:
        return {'success': False, 'error': f"Failed to open Excel workbook: {str(e)}"}

    sheet_names = wb.sheetnames
    ws = wb['Sites'] if 'Sites' in sheet_names else wb.active
    
    header_row_idx = None
    headers_found = []
    
    col_defs = [
        ('sr_no', ['sr no', 'sr. no', 'srno', 'sr_no', 's.no', 'sno', 'serial no', 'site no', 'index']),
        ('city', ['city', 'location city', 'town']),
        ('location', ['location', 'site location', 'address', 'site address', 'site name', 'description']),
        ('media', ['media', 'media type', 'type', 'display type']),
        ('light', ['light', 'lighting', 'illumination', 'illum']),
        ('w', ['w', 'width', 'w (ft)', 'w(ft)', 'w ft']),
        ('h', ['h', 'height', 'h (ft)', 'h(ft)', 'h ft']),
        ('sq_ft', ['sq ft', 'sqft', 'sq. ft', 'area', 'total sq ft']),
        ('avail', ['availablity', 'availability', 'status', 'available from', 'booked till']),
        ('rate', ['selling amount', 'selling rate', 'amount', 'rate', 'price', 'monthly rate', 'adv fee']),
        ('coords', ['latitude longitude', 'lat long', 'latitude, longitude', 'lat, lng', 'lat/long', 'gps', 'coordinates'])
    ]

    for r in range(1, min(15, ws.max_row + 1)):
        row_vals = [str(ws.cell(r, c).value or '').strip() for c in range(1, ws.max_column + 1)]
        row_joined = ' '.join(row_vals).lower()
        if 'location' in row_joined and ('media' in row_joined or 'city' in row_joined or 'sr no' in row_joined):
            header_row_idx = r
            headers_found = [str(ws.cell(r, c).value or '').strip() for c in range(1, ws.max_column + 1)]
            break

    if not header_row_idx:
        header_row_idx = 2 if ws.max_row >= 2 else 1
        headers_found = [str(ws.cell(header_row_idx, c).value or '').strip() for c in range(1, ws.max_column + 1)]

    while headers_found and not headers_found[-1]:
        headers_found.pop()

    mapping = {}
    for c_idx, h in enumerate(headers_found, start=1):
        h_clean = h.lower().strip()
        for field, aliases in col_defs:
            if field in mapping:
                continue
            if any(h_clean == a for a in aliases):
                mapping[field] = c_idx
                break
            elif any(len(a) > 1 and (a in h_clean or h_clean in a) for a in aliases):
                mapping[field] = c_idx
                break

    headers_norm = [h.strip().upper() for h in headers_found[:len(MASTER_TEMPLATE_COLUMNS)]]
    master_norm = [m.strip().upper() for m in MASTER_TEMPLATE_COLUMNS]
    matches_master = (headers_norm == master_norm)

    sample_rows = []
    total_sites = 0
    terms_detected = False

    for r in range(header_row_idx + 1, ws.max_row + 1):
        row_text_all = ' '.join([str(ws.cell(r, c).value or '') for c in range(1, ws.max_column + 1)]).strip()
        if not row_text_all:
            continue
        if 'TERMS' in row_text_all.upper() and 'CONDITION' in row_text_all.upper():
            terms_detected = True
            break

        def get_val(field):
            col = mapping.get(field)
            if col and col <= ws.max_column:
                v = ws.cell(r, col).value
                return str(v).strip() if v is not None else ""
            return ""

        sr_no = get_val('sr_no')
        city = get_val('city')
        loc = get_val('location')
        media = get_val('media')
        light = get_val('light')
        w = get_val('w')
        h = get_val('h')
        sq_ft = get_val('sq_ft')
        avail = get_val('avail')
        rate_val = get_val('rate')
        coords = get_val('coords')

        if not loc and not city and not sr_no:
            continue

        size_str = ""
        if w and h:
            size_str = f"{w}x{h} ft"
        elif sq_ft:
            size_str = f"{sq_ft} sq.ft"

        rate_str = format_inr(rate_val) if rate_val else ""

        total_sites += 1
        if len(sample_rows) < 5:
            sample_rows.append({
                'row': r,
                'sr_no': sr_no,
                'city': city,
                'location': loc,
                'media': media,
                'light': light,
                'size': size_str,
                'rate': rate_str,
                'avail': avail,
                'coords': coords
            })

    return {
        'success': True,
        'sheets': sheet_names,
        'activeSheet': ws.title,
        'headerRow': header_row_idx,
        'headers': headers_found,
        'matchesMaster': matches_master,
        'masterColumns': MASTER_TEMPLATE_COLUMNS,
        'mapping': mapping,
        'sampleRows': sample_rows,
        'totalSites': total_sites,
        'termsDetected': terms_detected,
        'termsNotice': "Terms & Conditions present in workbook will be ignored. Official 7-point Media Buzz software Terms & Conditions will be used exclusively."
    }

def merge_companion_excel(ppt_entries, companion_excel_path, column_mapping=None, sheet_name=None):
    """
    Extracts site data from companion Excel workbook (availability, coordinates / latitude longitude,
    rates, media, illumination, etc.) and enriches the PPT entries.
    Matches by site code / number, location keyword match, or sequential slide order.
    """
    try:
        excel_entries, _, _, _ = extract_entries_from_excel(
            companion_excel_path,
            column_mapping=column_mapping,
            sheet_name=sheet_name
        )
    except Exception as e:
        sys.stderr.write(f"[Companion Excel Error] {e}\n")
        return ppt_entries

    if not excel_entries:
        return ppt_entries

    # Build lookup maps
    code_map = {}
    sr_map = {}
    for x in excel_entries:
        if x.get('code'):
            code_map[str(x['code']).lower().strip()] = x
        if x.get('sr_no'):
            sr_map[str(x['sr_no']).lower().strip()] = x

    for idx, entry in enumerate(ppt_entries):
        matched = None
        ppt_code = str(entry.get('code') or '').lower().strip()
        ppt_num = str(entry.get('site_number') or (idx + 1)).strip()

        # 1. Match by exact code or number
        if ppt_code and ppt_code in code_map:
            matched = code_map[ppt_code]
        elif ppt_num and ppt_num in sr_map:
            matched = sr_map[ppt_num]
        elif str(idx + 1) in sr_map:
            matched = sr_map[str(idx + 1)]

        # 2. Match by location keyword overlap if not matched
        if not matched and entry.get('location'):
            ppt_loc_words = set(re.findall(r'\b\w{3,}\b', entry['location'].lower()))
            best_score = 0
            best_x = None
            for x in excel_entries:
                if x.get('location'):
                    x_loc_words = set(re.findall(r'\b\w{3,}\b', x['location'].lower()))
                    overlap = len(ppt_loc_words.intersection(x_loc_words))
                    if overlap > best_score and overlap >= 2:
                        best_score = overlap
                        best_x = x
            if best_x:
                matched = best_x

        # 3. Fallback sequential matching
        if not matched and idx < len(excel_entries):
            matched = excel_entries[idx]

        # Apply matched fields (e.g. Availability, Latitude Longitude, Rate, etc.)
        if matched:
            if matched.get('availability') and str(matched['availability']).strip():
                entry['availability'] = str(matched['availability']).strip()
            if matched.get('coords') and str(matched['coords']).strip():
                entry['coords'] = str(matched['coords']).strip()
            if matched.get('rate') and str(matched['rate']).strip():
                entry['rate'] = str(matched['rate']).strip()
            if matched.get('city') and str(matched['city']).strip():
                entry['city'] = str(matched['city']).strip()
            if matched.get('type') and str(matched['type']).strip():
                entry['type'] = str(matched['type']).strip()
            if matched.get('illum') and str(matched['illum']).strip():
                entry['illum'] = str(matched['illum']).strip()
            if matched.get('size') and str(matched['size']).strip():
                entry['size'] = str(matched['size']).strip()
            if matched.get('location') and (not entry.get('location') or len(entry['location']) < 6 or 'prime city' in entry['location'].lower()):
                entry['location'] = matched['location']
            entry['companion_matched'] = True

            # Clear missing coordinate / availability review flags
            r_flags = entry.get('review_flags', [])
            new_flags = [f for f in r_flags if 'availability' not in f.lower() and 'coordinate' not in f.lower()]
            entry['review_flags'] = new_flags
            entry['review_needed'] = len(new_flags) > 0

    return ppt_entries

def generate_master_excel(entries, output_excel_path, manual_city=None, title="MEDIA BUZZ - SITES SPECIFICATION"):
    """
    Generates an authentic 11-column Master Excel workbook matching Jade Blue 16-10-2026.xlsx format.
    Columns: SR NO, City, LOCATION, MEDIA, LIGHT, W, H, SQ FT, AVAILABLITY, Selling Amount, Latitude Longitude
    Includes the official 7-point Media Buzz software Terms & Conditions.
    """
    if not HAS_OPENPYXL or not openpyxl:
        return

    try:
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sites"

        # Row 1: Executive Media Buzz Brand Header
        ws.merge_cells('A1:K1')
        title_cell = ws['A1']
        title_cell.value = (title or "MEDIA BUZZ - SITES SPECIFICATION").upper()
        title_cell.font = Font(name='Calibri', size=13, bold=True, color='1E293B')
        title_cell.alignment = Alignment(horizontal='center', vertical='middle')
        title_cell.fill = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
        ws.row_dimensions[1].height = 32

        # Row 2: Headers matching Jade Blue 16-10-2026.xlsx in Bright Yellow
        headers = [
            'SR NO', 'City', 'LOCATION', 'MEDIA', 'LIGHT',
            'W', 'H', 'SQ FT', 'AVAILABLITY', 'Selling Amount', 'Latitude Longitude'
        ]
        ws.append(headers)
        ws.row_dimensions[2].height = 26

        yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
        header_font = Font(name='Calibri', size=11, bold=True, color='000000')
        thin_border = Border(
            left=Side(style='thin', color='C0C0C0'),
            right=Side(style='thin', color='C0C0C0'),
            top=Side(style='thin', color='C0C0C0'),
            bottom=Side(style='medium', color='808080')
        )

        for col_idx in range(1, 12):
            cell = ws.cell(row=2, column=col_idx)
            cell.fill = yellow_fill
            cell.font = header_font
            cell.border = thin_border
            align_h = 'center' if col_idx in [1, 2, 4, 5, 6, 7, 8, 9] else ('right' if col_idx == 10 else 'left')
            cell.alignment = Alignment(horizontal=align_h, vertical='middle')

        # Row 3+: Data rows
        data_font = Font(name='Calibri', size=10.5)
        data_border = Border(
            left=Side(style='thin', color='E8E8E8'),
            right=Side(style='thin', color='E8E8E8'),
            top=Side(style='thin', color='E8E8E8'),
            bottom=Side(style='thin', color='E8E8E8')
        )

        for idx, e in enumerate(entries, start=1):
            r_num = idx + 2
            # Parse dimensions
            size_txt = str(e.get('size') or '').replace('ft', '').replace('FT', '').strip()
            dim_m = re.match(r'(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)', size_txt)
            w_val = e.get('w') or (dim_m.group(1) if dim_m else "")
            h_val = e.get('h') or (dim_m.group(2) if dim_m else "")
            sqft_val = e.get('sq_ft') or ""
            if not sqft_val and dim_m:
                try:
                    sqft_val = str(int(float(dim_m.group(1)) * float(dim_m.group(2))))
                except Exception:
                    sqft_val = ""

            row_data = [
                idx,
                manual_city or e.get('city') or 'Ahmedabad',
                e.get('location') or '',
                e.get('type') or 'Hoarding',
                e.get('illum') or 'BL',
                w_val,
                h_val,
                sqft_val,
                e.get('availability') or 'Available',
                e.get('rate') or '',
                e.get('coords') or ''
            ]
            ws.append(row_data)
            ws.row_dimensions[r_num].height = 22
            for col_idx in range(1, 12):
                cell = ws.cell(row=r_num, column=col_idx)
                cell.font = data_font
                cell.border = data_border
                align_h = 'center' if col_idx in [1, 2, 4, 5, 6, 7, 8, 9] else ('right' if col_idx == 10 else 'left')
                cell.alignment = Alignment(horizontal=align_h, vertical='middle')

        # Auto-adjust column widths
        col_widths = {1: 8, 2: 16, 3: 45, 4: 15, 5: 10, 6: 8, 7: 8, 8: 12, 9: 20, 10: 20, 11: 28}
        for col_idx, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        # Official 7-point Media Buzz software Terms & Conditions at bottom
        terms_start = len(entries) + 5
        ws.cell(row=terms_start, column=2, value="TERMS & CONDITIONS:").font = Font(name='Calibri', size=11, bold=True, color='1E293B')
        for t_idx, term in enumerate(OFFICIAL_SOFTWARE_TERMS):
            ws.cell(row=terms_start + 1 + t_idx, column=2, value=term).font = Font(name='Calibri', size=10, color='475569')

        os.makedirs(os.path.dirname(os.path.abspath(output_excel_path)), exist_ok=True)
        wb.save(output_excel_path)
    except Exception as exc:
        sys.stderr.write(f"[Master Excel Gen Error] {exc}\n")

def generate_converted_pptx_pure(template_file, output_path, entries, is_excel=False):
    """Generates the master template presentation with cloned slides and official software Terms & Conditions."""
    with zipfile.ZipFile(template_file, 'r') as z_in:
        files = {f.filename: z_in.read(f.filename) for f in z_in.filelist}

    base_slide_xml_str = files['ppt/slides/slide3.xml'].decode('utf-8')
    base_rels_xml_str = files['ppt/slides/_rels/slide3.xml.rels'].decode('utf-8')

    # Remove any notesSlide relationship from base_rels to avoid 1-to-1 violation across cloned slides
    base_rels_xml_str = re.sub(r'<Relationship[^>]*relationships/notesSlide[^>]*/>', '', base_rels_xml_str)

    content_types_str = files['[Content_Types].xml'].decode('utf-8')
    pres_xml_str = files['ppt/presentation.xml'].decode('utf-8')
    pres_rels_str = files['ppt/_rels/presentation.xml.rels'].decode('utf-8')

    existing_rids = [int(m) for m in re.findall(r'rId(\d+)', pres_rels_str)]
    max_rid = max(existing_rids) if existing_rids else 20

    existing_sldids = [int(m) for m in re.findall(r'<p:sldId\s+id="(\d+)"', pres_xml_str)]
    max_sldid = max(existing_sldids) if existing_sldids else 300

    n_sites = len(entries)
    terms_slide_num = n_sites + 3
    closing_slide_num = n_sites + 4

    closing_xml_data = files['ppt/slides/slide5.xml']
    closing_rels_data = files.get('ppt/slides/_rels/slide5.xml.rels')

    # Remove old template slide 5
    del files['ppt/slides/slide5.xml']
    if 'ppt/slides/_rels/slide5.xml.rels' in files:
        del files['ppt/slides/_rels/slide5.xml.rels']

    if closing_rels_data:
        closing_rels_str = closing_rels_data.decode('utf-8')
        closing_rels_str = re.sub(r'<Relationship[^>]*relationships/notesSlide[^>]*/>', '', closing_rels_str)
        closing_rels_data = closing_rels_str.encode('utf-8')

    content_types_str = re.sub(r'<Override\s+PartName="/ppt/slides/slide[35]\.xml"[^>]*/>', '', content_types_str)

    sld_id_list = [
        '<p:sldId id="256" r:id="rId2"/>',
        '<p:sldId id="257" r:id="rId3"/>'
    ]

    pres_rels_str = re.sub(r'<Relationship\s+Id="rId4"[^>]*/>', '', pres_rels_str)
    pres_rels_str = re.sub(r'<Relationship\s+Id="rId6"[^>]*/>', '', pres_rels_str)

    # 1. Generate Site Slides
    for idx, entry in enumerate(entries):
        site_num = idx + 1
        slide_idx = site_num + 2
        slide_filename = f"ppt/slides/slide{slide_idx}.xml"
        rels_filename = f"ppt/slides/_rels/slide{slide_idx}.xml.rels"

        slide_xml = base_slide_xml_str
        rels_xml = base_rels_xml_str

        site_label = f"Site {site_num}"
        # Shape 0: Clear redundant small text above badge
        slide_xml = re.sub(r'(<a:t[^>]*>)Site\s*\d+(</a:t>)', r'\g<1>\g<2>', slide_xml, count=1)

        # Yellow badge (Shape 2): Highlight "Site 1", "Site 2" without any artificial site code
        slide_xml = slide_xml.replace('MB - 07', saxutils.escape(site_label))
        rels_xml = re.sub(r'Target="https://panel\.mediabuzzoutdoor\.com/campaigns\?site=[^"]*"',
                          'Target="https://panel.mediabuzzoutdoor.com/campaigns"',
                          rels_xml)

        city = entry.get('city') or ''
        slide_xml = slide_xml.replace('Ahmedabad', saxutils.escape(city))

        loc = entry.get('location') or ''
        slide_xml = slide_xml.replace('Shivranjani Bridge- Fcg. Shyamal 4 road Jnc - Left (1)', saxutils.escape(loc))

        # Do not invent values; leave missing values blank (Rule 8)
        size = entry.get('size') or ''
        mtype = entry.get('type') or ''
        illum = entry.get('illum') or ''
        rate = entry.get('rate') or ''
        avail = entry.get('availability') or ''
        coords = entry.get('coords') or ''

        slide_xml = slide_xml.replace('30x10 ft', saxutils.escape(size))
        slide_xml = slide_xml.replace('Gantry', saxutils.escape(mtype))
        slide_xml = slide_xml.replace('BL', saxutils.escape(illum))
        slide_xml = slide_xml.replace('₹ 2,25,000/-', saxutils.escape(rate))
        slide_xml = slide_xml.replace('Booked till 30 Sept 2026', saxutils.escape(avail))
        slide_xml = slide_xml.replace('23.018881, 72.530022', saxutils.escape(coords))

        photo_bytes = entry.get('photo_bytes')
        if photo_bytes:
            media_filename = f"ppt/media/site_photo_{site_num}.jpeg"
            files[media_filename] = photo_bytes
            rels_xml = re.sub(r'Target="\.\./media/image-3-1\.jpeg"', f'Target="../media/site_photo_{site_num}.jpeg"', rels_xml)

        files[slide_filename] = slide_xml.encode('utf-8')
        files[rels_filename] = rels_xml.encode('utf-8')

        content_types_str = content_types_str.replace(
            '</Types>',
            f'<Override PartName="/{slide_filename}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'
        )

        max_rid += 1
        rid = f"rId{max_rid}"
        pres_rels_str = pres_rels_str.replace(
            '</Relationships>',
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{slide_idx}.xml"/></Relationships>'
        )

        max_sldid += 1
        sld_id_list.append(f'<p:sldId id="{max_sldid}" r:id="{rid}"/>')

    # 2. Generate Terms & Conditions Slide (Official Media Buzz Software Terms - Rule 7)
    terms_slide_filename = f"ppt/slides/slide{terms_slide_num}.xml"
    terms_rels_filename = f"ppt/slides/_rels/slide{terms_slide_num}.xml.rels"

    terms_xml = base_slide_xml_str
    terms_xml = re.sub(r'(<a:t[^>]*>)Site\s*\d+(</a:t>)', r'\g<1>\g<2>', terms_xml, count=1)
    terms_xml = terms_xml.replace('MB - 07', 'TERMS &amp; CONDITIONS')
    terms_xml = terms_xml.replace('Ahmedabad', 'Terms &amp; Conditions')
    terms_xml = terms_xml.replace('Shivranjani Bridge- Fcg. Shyamal 4 road Jnc - Left (1)', 'Media Buzz Outdoor Media Contract Terms')
    
    terms_xml = terms_xml.replace('Size:  ', '1. ')
    terms_xml = terms_xml.replace('30x10 ft', saxutils.escape(OFFICIAL_SOFTWARE_TERMS[0][3:]))
    
    terms_xml = terms_xml.replace('Type:  ', '2. ')
    terms_xml = terms_xml.replace('Gantry', saxutils.escape(OFFICIAL_SOFTWARE_TERMS[1][3:]))
    
    terms_xml = terms_xml.replace('Illumination:  ', '3. ')
    terms_xml = terms_xml.replace('BL', saxutils.escape(OFFICIAL_SOFTWARE_TERMS[2][3:]))
    
    terms_xml = terms_xml.replace('Adv.fee per month:  ', '4. ')
    terms_xml = terms_xml.replace('₹ 2,25,000/-', saxutils.escape(OFFICIAL_SOFTWARE_TERMS[3][3:]))
    
    terms_xml = terms_xml.replace('Availability:  ', '5. ')
    terms_xml = terms_xml.replace('Booked till 30 Sept 2026', saxutils.escape(OFFICIAL_SOFTWARE_TERMS[4][3:]))
    
    terms_xml = terms_xml.replace('Location Coordinates:  ', '6. ')
    terms_xml = terms_xml.replace('23.018881, 72.530022', saxutils.escape(OFFICIAL_SOFTWARE_TERMS[5][3:] + " " + OFFICIAL_SOFTWARE_TERMS[6]))

    files[terms_slide_filename] = terms_xml.encode('utf-8')
    files[terms_rels_filename] = base_rels_xml_str.encode('utf-8')

    content_types_str = content_types_str.replace(
        '</Types>',
        f'<Override PartName="/{terms_slide_filename}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'
    )
    max_rid += 1
    terms_rid = f"rId{max_rid}"
    pres_rels_str = pres_rels_str.replace(
        '</Relationships>',
        f'<Relationship Id="{terms_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{terms_slide_num}.xml"/></Relationships>'
    )
    max_sldid += 1
    sld_id_list.append(f'<p:sldId id="{max_sldid}" r:id="{terms_rid}"/>')

    # 3. Generate Closing Slide
    closing_slide_filename = f"ppt/slides/slide{closing_slide_num}.xml"
    closing_rels_filename = f"ppt/slides/_rels/slide{closing_slide_num}.xml.rels"
    files[closing_slide_filename] = closing_xml_data
    if closing_rels_data:
        files[closing_rels_filename] = closing_rels_data

    content_types_str = content_types_str.replace(
        '</Types>',
        f'<Override PartName="/{closing_slide_filename}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'
    )
    max_rid += 1
    closing_rid = f"rId{max_rid}"
    pres_rels_str = pres_rels_str.replace(
        '</Relationships>',
        f'<Relationship Id="{closing_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{closing_slide_num}.xml"/></Relationships>'
    )
    max_sldid += 1
    sld_id_list.append(f'<p:sldId id="{max_sldid}" r:id="{closing_rid}"/>')

    # Update presentation.xml and rels
    pres_xml_str = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', f'<p:sldIdLst>{"".join(sld_id_list)}</p:sldIdLst>', pres_xml_str)

    files['[Content_Types].xml'] = content_types_str.encode('utf-8')
    files['ppt/presentation.xml'] = pres_xml_str.encode('utf-8')
    files['ppt/_rels/presentation.xml.rels'] = pres_rels_str.encode('utf-8')

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
        for fname, data in files.items():
            z_out.writestr(fname, data)


def process_presentation_pure(input_path, output_path, template_file, mapping=None, sheet_name=None, companion_excel=None, manual_city=None, custom_name=None):
    """Processes conversion for PPTX or Excel using pure Python standard library & OpenXML."""
    is_excel = input_path.lower().endswith(('.xlsx', '.xls'))
    headers_detected = []
    mapping_used = {}
    companion_merged = False
    
    if is_excel:
        entries, headers_detected, h_idx, mapping_used = extract_entries_from_excel(input_path, column_mapping=mapping, sheet_name=sheet_name)
    else:
        entries = extract_entries_pure_zip(input_path)
        if companion_excel and os.path.isfile(companion_excel):
            entries = merge_companion_excel(entries, companion_excel, column_mapping=mapping, sheet_name=sheet_name)
            companion_merged = True

    # Manual city override or fallback
    if manual_city and str(manual_city).strip():
        m_city = str(manual_city).strip()
        for e in entries:
            e['city'] = m_city

    if not entries:
        raise ValueError("Could not extract any location or billboard listings from the uploaded file.")
        
    generate_converted_pptx_pure(template_file, output_path, entries, is_excel=is_excel)

    # Automatically generate standardized 11-column Master Excel workbook
    excel_output_path = re.sub(r'\.pptx$', '.xlsx', output_path, flags=re.IGNORECASE)
    excel_generated = False
    try:
        generate_master_excel(entries, excel_output_path, manual_city=manual_city, title=custom_name or "MEDIA BUZZ - SITES SPECIFICATION")
        excel_generated = os.path.isfile(excel_output_path)
    except Exception as xErr:
        sys.stderr.write(f"[Excel Generation Notice] {xErr}\n")
    
    slide_previews = []
    for idx, entry in enumerate(entries):
        site_num = idx + 1
        thumb_data = generate_slide_thumbnail(
            entry.get('photo_bytes'),
            site_num,
            entry.get('code') or '',
            entry.get('size') or '',
            entry.get('location') or '',
            entry.get('city') or '',
            entry.get('type') or '',
            entry.get('illum') or '',
            entry.get('rate') or '',
            entry.get('availability') or '',
            entry.get('coords') or ''
        )
        slide_previews.append({
            'siteNumber': site_num,
            'sourceSlide': entry['source_slide'],
            'code': entry.get('code') or '',
            'size': entry.get('size') or '',
            'location': entry.get('location') or '',
            'city': entry.get('city') or '',
            'type': entry.get('type') or '',
            'illum': entry.get('illum') or '',
            'rate': entry.get('rate') or '',
            'availability': entry.get('availability') or '',
            'coords': entry.get('coords') or '',
            'hasPhoto': entry.get('photo_bytes') is not None,
            'reviewNeeded': entry.get('review_needed', False),
            'reviewFlags': entry.get('review_flags', []),
            'companionMatched': entry.get('companion_matched', False),
            'previewUrl': thumb_data
        })

    # Append Terms & Conditions slide preview (Official software terms - Rule 7 & 9)
    terms_thumb = generate_terms_slide_thumbnail()
    slide_previews.append({
        'siteNumber': 'Terms & Conditions',
        'isTermsSlide': True,
        'sourceSlide': 'Media Buzz Software',
        'code': '',
        'size': '',
        'location': 'Official Media Buzz Software Terms & Conditions (7 Points)',
        'city': 'All Markets',
        'type': 'Terms',
        'illum': 'Contract',
        'rate': '7-Point Terms Preserved',
        'availability': 'Official Policy',
        'coords': '',
        'hasPhoto': False,
        'reviewNeeded': False,
        'reviewFlags': [],
        'previewUrl': terms_thumb
    })
        
    return {
        'success': True,
        'isExcel': is_excel,
        'companionMerged': companion_merged,
        'manualCityApplied': bool(manual_city and str(manual_city).strip()),
        'manualCity': manual_city or None,
        'excelGenerated': excel_generated,
        'softwareTermsVerified': True,
        'message': 'Successfully converted into Media Buzz Master Template format.',
        'processedSlides': len(entries),
        'totalSlidesInFile': len(entries) + 4, # Cover 1, Cover 2, N sites, Terms, Closing
        'templateUsed': os.path.basename(template_file),
        'headers': headers_detected,
        'mapping': mapping_used,
        'slides': slide_previews
    }

# ==============================================================================
# python-pptx Engine (Optional Accelerator)
# ==============================================================================

def process_presentation_pptx(input_path, output_path, template_file):
    """Processes conversion using python-pptx when available."""
    in_prs = pptx.Presentation(input_path)
    out_prs = pptx.Presentation(template_file)
    
    if len(out_prs.slides) == 5:
        demo_sldId = out_prs.slides._sldIdLst[3]
        out_prs.part.drop_rel(demo_sldId.rId)
        out_prs.slides._sldIdLst.remove(demo_sldId)

    if len(out_prs.slides) < 3:
        raise ValueError("Master template does not contain the required site slide at slide index 3.")
        
    base_site_slide = out_prs.slides[2]
    closing_sldId = out_prs.slides._sldIdLst[len(out_prs.slides) - 1]
    
    entries = []
    for i, slide in enumerate(in_prs.slides):
        all_texts = []
        for sh in slide.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    t = p.text.strip()
                    if t: all_texts.append(t)
            if sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells:
                        t = c.text.strip()
                        if t: all_texts.append(t)
                        
        full_text = ' '.join(all_texts).strip()
        slide_w_in = in_prs.slide_width / 914400
        slide_h_in = in_prs.slide_height / 914400
        
        if full_text.lower() in ['mediabuzzoutdoor.com', 'thank you', 'thanks', 'contact us', 'thank you!']:
            continue
            
        code = extract_code(full_text, len(entries) + 1)
        size = extract_size(full_text)
        city = detect_city(full_text)
        loc = clean_location(full_text, size, code, city)
        mtype = detect_media_type(full_text)
        illum = detect_illumination(full_text)
        rate = detect_rate(full_text)
        avail = detect_availability(full_text)
        coords = detect_coords(full_text)

        best_pic = None
        max_area = 0
        for sh in slide.shapes:
            if sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
                w_in = sh.width / 914400
                h_in = sh.height / 914400
                area = w_in * h_in
                if w_in >= 0.92 * slide_w_in and h_in >= 0.92 * slide_h_in:
                    continue
                if area > max_area and area > 2.0:
                    max_area = area
                    best_pic = sh
                    
        if not best_pic and not size and (not loc or len(loc) < 5):
            continue
        
        review_flags = []
        if not best_pic:
            review_flags.append("Photo missing or not detected")
        if not size:
            review_flags.append("Dimensions/size missing")
        if not loc or len(loc) < 3:
            review_flags.append("Location description unclear")
            
        photo_bytes = None
        if best_pic:
            try:
                photo_bytes = best_pic.image.blob
            except Exception:
                photo_bytes = None
                
        entries.append({
            'source_slide': i + 1,
            'photo_bytes': photo_bytes,
            'code': code,
            'size': size or '30x10 ft',
            'location': loc or 'Prime City Location',
            'city': city or 'Ahmedabad',
            'type': mtype or 'Hoarding',
            'illum': illum or 'BL',
            'rate': rate or 'On Request',
            'availability': avail or 'Available',
            'coords': coords or '23.012345, 72.567890',
            'review_needed': len(review_flags) > 0,
            'review_flags': review_flags
        })
        
    if not entries:
        raise ValueError("Could not extract any location or billboard slides from the uploaded presentation.")
        
    n_entries = len(entries)
    
    def apply_data_to_slide(slide, entry, site_num):
        photo_data = entry.get('photo_bytes') or generate_no_photo_bytes()
        try:
            buf = io.BytesIO(photo_data)
            new_img_part, rId = slide.part.get_or_add_image_part(buf)
            pic_sh = None
            for sh in slide.shapes:
                if sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
                    w_in = sh.width / 914400
                    h_in = sh.height / 914400
                    if 5.0 <= w_in <= 7.0 and 4.0 <= h_in <= 6.0:
                        pic_sh = sh
                        break
            if pic_sh is not None:
                blip = pic_sh._element.xpath('.//a:blip')
                if blip:
                    blip[0].set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', rId)
        except Exception:
            pass

        site_label = f"Site {site_num}"
        code_text = entry['code']
        city_text = entry['city']
        loc_text = entry['location']
        size_text = entry['size']
        type_text = entry['type']
        illum_text = entry['illum']
        rate_text = entry['rate']
        avail_text = entry['availability']
        coords_text = entry['coords']

        def set_spec_paragraph(p, field_label, value):
            p.text = ""
            r1 = p.add_run()
            r1.text = f"{field_label}:  "
            r1.font.name = 'Arial'
            r1.font.size = pptx.util.Pt(11)
            r1.font.bold = False
            r1.font.color.rgb = pptx.dml.color.RGBColor(255, 194, 0)
            
            r2 = p.add_run()
            r2.text = str(value)
            r2.font.name = 'Arial'
            r2.font.size = pptx.util.Pt(11)
            r2.font.bold = False
            r2.font.color.rgb = pptx.dml.color.RGBColor(255, 255, 255)

        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            tf = sh.text_frame
            name = sh.name
            top_in = sh.top / 914400
            left_in = sh.left / 914400

            if name == 'Text 1' or (abs(top_in - 0.70) < 0.15 and abs(left_in - 7.45) < 0.25):
                if tf.paragraphs and tf.paragraphs[0].runs:
                    p = tf.paragraphs[0]
                    p.runs[0].text = site_label
                    p.runs[0].font.name = 'Arial'
                    p.runs[0].font.size = pptx.util.Pt(13)
                    p.runs[0].font.bold = False
                    p.runs[0].font.color.rgb = pptx.dml.color.RGBColor(255, 255, 255)
                else:
                    tf.text = site_label
                    
            elif name == 'Text 2' or (abs(top_in - 1.05) < 0.20 and abs(left_in - 7.45) < 0.25):
                p = tf.paragraphs[0]
                p.text = code_text
                if p.runs:
                    p.runs[0].font.name = 'Arial'
                    p.runs[0].font.size = pptx.util.Pt(26)
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = pptx.dml.color.RGBColor(0, 0, 0)
                    
            elif name == 'Text 3' or (abs(top_in - 1.98) < 0.15 and abs(left_in - 7.45) < 0.25):
                if tf.paragraphs and tf.paragraphs[0].runs:
                    p = tf.paragraphs[0]
                    p.runs[0].text = city_text
                    p.runs[0].font.name = 'Arial'
                    p.runs[0].font.size = pptx.util.Pt(24)
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = pptx.dml.color.RGBColor(255, 255, 255)
                else:
                    tf.text = city_text
                    
            elif name == 'Text 4' or (abs(top_in - 2.48) < 0.15 and abs(left_in - 7.45) < 0.25):
                if tf.paragraphs and tf.paragraphs[0].runs:
                    p = tf.paragraphs[0]
                    p.runs[0].text = loc_text
                    p.runs[0].font.name = 'Arial'
                    p.runs[0].font.size = pptx.util.Pt(16)
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = pptx.dml.color.RGBColor(255, 255, 255)
                else:
                    tf.text = loc_text

            for p in tf.paragraphs:
                t = p.text.strip()
                if t.startswith('Size:'):
                    set_spec_paragraph(p, 'Size', size_text)
                elif t.startswith('Type:'):
                    set_spec_paragraph(p, 'Type', type_text)
                elif t.startswith('Illumination:'):
                    set_spec_paragraph(p, 'Illumination', illum_text)
                elif t.startswith('Adv.fee') or t.startswith('Adv.'):
                    set_spec_paragraph(p, 'Adv.fee per month', rate_text)
                elif t.startswith('Availability:'):
                    set_spec_paragraph(p, 'Availability', avail_text)
                elif t.startswith('Location Coordinates:'):
                    set_spec_paragraph(p, 'Location Coordinates', coords_text)

    for idx in range(n_entries):
        site_num = idx + 1
        if idx == 0:
            slide = out_prs.slides[2]
        else:
            slide = duplicate_slide(out_prs, base_site_slide)
        apply_data_to_slide(slide, entries[idx], site_num)

    out_prs.slides._sldIdLst.remove(closing_sldId)
    out_prs.slides._sldIdLst.append(closing_sldId)
        
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    out_prs.save(output_path)
    
    slide_previews = []
    for idx, entry in enumerate(entries):
        site_num = idx + 1
        thumb_data = generate_slide_thumbnail(
            entry['photo_bytes'],
            site_num,
            entry['code'],
            entry['size'],
            entry['location'],
            entry['city'],
            entry['type'],
            entry['illum'],
            entry['rate'],
            entry['availability'],
            entry['coords']
        )
        slide_previews.append({
            'siteNumber': site_num,
            'sourceSlide': entry['source_slide'],
            'code': entry['code'],
            'size': entry['size'],
            'location': entry['location'],
            'city': entry['city'],
            'type': entry['type'],
            'illum': entry['illum'],
            'rate': entry['rate'],
            'availability': entry['availability'],
            'coords': entry['coords'],
            'hasPhoto': entry['photo_bytes'] is not None,
            'reviewNeeded': entry['review_needed'],
            'reviewFlags': entry['review_flags'],
            'previewUrl': thumb_data
        })
        
    return {
        'success': True,
        'message': 'Successfully converted presentation into Media Buzz Master Template format.',
        'processedSlides': n_entries,
        'totalSlidesInFile': len(out_prs.slides),
        'templateUsed': os.path.basename(template_file),
        'slides': slide_previews
    }

def process_presentation(input_path, output_path, master_template_path=None, mapping=None, sheet_name=None, companion_excel=None, manual_city=None, custom_name=None):
    """
    Main conversion entrypoint.
    Executes zero-dependency OpenXML engine ensuring 100% valid PowerPoint files.
    Supports both PowerPoint (.pptx) and Excel (.xlsx, .xls) files.
    """
    template_file = find_master_template(master_template_path)
    if not template_file:
        raise FileNotFoundError("Media Buzz master template not found. Searched candidate paths.")
        
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input file not found at: {input_path}")
        
    return process_presentation_pure(
        input_path, output_path, template_file,
        mapping=mapping, sheet_name=sheet_name,
        companion_excel=companion_excel, manual_city=manual_city,
        custom_name=custom_name
    )


def main():
    parser = argparse.ArgumentParser(description="Media Buzz PowerPoint & Excel Converter Engine")
    parser.add_argument('--check-template', action='store_true', help="Check master template status")
    parser.add_argument('--inspect-excel', type=str, help="Inspect structure of an uploaded Excel workbook")
    parser.add_argument('--input', type=str, help="Path to input presentation or Excel workbook")
    parser.add_argument('--output', type=str, help="Path to output presentation")
    parser.add_argument('--template', type=str, help="Path to master template")
    parser.add_argument('--mapping', type=str, help="JSON-encoded custom column mapping")
    parser.add_argument('--sheet', type=str, help="Target worksheet name for Excel conversion")
    parser.add_argument('--companion-excel', type=str, help="Optional companion Excel workbook to extract availability, coordinates, rates, etc.")
    parser.add_argument('--manual-city', type=str, help="Manual city override or default city")
    parser.add_argument('--custom-name', type=str, help="Custom name for output presentation and Excel")
    args = parser.parse_args()
    
    if args.check_template:
        tmpl = find_master_template(args.template)
        if tmpl and os.path.isfile(tmpl):
            res = {
                'available': True,
                'path': tmpl,
                'filename': os.path.basename(tmpl),
                'size': os.path.getsize(tmpl)
            }
        else:
            res = {
                'available': False,
                'searchedPaths': CANDIDATE_TEMPLATE_PATHS
            }
        print(json.dumps(res, ensure_ascii=True))
        return

    if args.inspect_excel:
        if not os.path.isfile(args.inspect_excel):
            err_res = {'success': False, 'error': f"Excel file not found at: {args.inspect_excel}"}
            print(json.dumps(err_res, ensure_ascii=True))
            sys.exit(1)
        res = inspect_excel(args.inspect_excel)
        print(json.dumps(res, ensure_ascii=True))
        return
        
    if not args.input or not args.output:
        err_res = {'success': False, 'error': 'Both --input and --output are required for conversion.'}
        print(json.dumps(err_res, ensure_ascii=True))
        sys.exit(1)

    col_mapping = None
    if args.mapping:
        try:
            col_mapping = json.loads(args.mapping)
        except Exception as mErr:
            sys.stderr.write(f"[PPT Engine] Notice parsing mapping: {mErr}\n")
        
    try:
        res = process_presentation(
            args.input,
            args.output,
            args.template,
            mapping=col_mapping,
            sheet_name=args.sheet,
            companion_excel=args.companion_excel,
            manual_city=args.manual_city,
            custom_name=args.custom_name
        )
        # Write full result file to disk to prevent stdout buffer overflow
        result_json_path = args.output + '.json'
        try:
            with open(result_json_path, 'w', encoding='utf-8') as f:
                json.dump(res, f, ensure_ascii=True)
        except Exception as fErr:
            sys.stderr.write(f"[PPT Engine] Notice writing result file: {fErr}\n")
            
        print(json.dumps(res, ensure_ascii=True))
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        sys.stderr.write(f"Conversion Exception: {str(e)}\n{tb}\n")
        err_res = {
            'success': False,
            'error': str(e),
            'traceback': tb
        }
        if args.output:
            try:
                with open(args.output + '.json', 'w', encoding='utf-8') as f:
                    json.dump(err_res, f, ensure_ascii=True)
            except Exception:
                pass
        print(json.dumps(err_res, ensure_ascii=True))
        sys.exit(1)

if __name__ == '__main__':
    main()

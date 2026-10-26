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

# Optional pptx import (graceful fallback if not installed)
try:
    import pptx
except ImportError:
    pptx = None


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
            
            draw.text((rx, 118), city or "Ahmedabad", fill=(255, 255, 255))
            draw.text((rx, 148), location[:44] if location else "Location Details", fill=(255, 255, 255))
            if len(location) > 44:
                draw.text((rx, 168), location[44:88], fill=(255, 255, 255))
                
            fields = [
                ("Size:", size or "30x10 ft"),
                ("Type:", mtype or "Hoarding"),
                ("Illumination:", illum or "BL"),
                ("Adv.fee per month:", rate or "On Request"),
                ("Availability:", avail or "Available"),
                ("Location Coordinates:", coords or "23.012345, 72.567890")
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

def generate_converted_pptx_pure(template_file, output_path, entries):
    """Generates the master template presentation with cloned slides using zipfile only."""
    with zipfile.ZipFile(template_file, 'r') as z_in:
        files = {f.filename: z_in.read(f.filename) for f in z_in.filelist}

    base_slide_xml_str = files['ppt/slides/slide3.xml'].decode('utf-8')
    base_rels_xml_str = files['ppt/slides/_rels/slide3.xml.rels'].decode('utf-8')

    # Remove any notesSlide relationship from base_rels to avoid 1-to-1 violation across cloned slides
    base_rels_xml_str = re.sub(r'<Relationship[^>]*relationships/notesSlide[^>]*/>', '', base_rels_xml_str)

    content_types_str = files['[Content_Types].xml'].decode('utf-8')
    pres_xml_str = files['ppt/presentation.xml'].decode('utf-8')
    pres_rels_str = files['ppt/_rels/presentation.xml.rels'].decode('utf-8')

    # Find highest rId in presentation.xml.rels
    existing_rids = [int(m) for m in re.findall(r'rId(\d+)', pres_rels_str)]
    max_rid = max(existing_rids) if existing_rids else 20

    # ONLY match sldId id attributes, NOT sldMasterId id="2147483648"!
    existing_sldids = [int(m) for m in re.findall(r'<p:sldId\s+id="(\d+)"', pres_xml_str)]
    max_sldid = max(existing_sldids) if existing_sldids else 300

    n_sites = len(entries)
    closing_slide_num = n_sites + 3
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

    closing_slide_filename = f"ppt/slides/slide{closing_slide_num}.xml"
    closing_rels_filename = f"ppt/slides/_rels/slide{closing_slide_num}.xml.rels"
    files[closing_slide_filename] = closing_xml_data
    if closing_rels_data:
        files[closing_rels_filename] = closing_rels_data

    content_types_str = re.sub(r'<Override\s+PartName="/ppt/slides/slide[35]\.xml"[^>]*/>', '', content_types_str)

    sld_id_list = [
        '<p:sldId id="256" r:id="rId2"/>',
        '<p:sldId id="257" r:id="rId3"/>'
    ]

    pres_rels_str = re.sub(r'<Relationship\s+Id="rId4"[^>]*/>', '', pres_rels_str)
    pres_rels_str = re.sub(r'<Relationship\s+Id="rId6"[^>]*/>', '', pres_rels_str)

    for idx, entry in enumerate(entries):
        site_num = idx + 1
        slide_idx = site_num + 2
        slide_filename = f"ppt/slides/slide{slide_idx}.xml"
        rels_filename = f"ppt/slides/_rels/slide{slide_idx}.xml.rels"

        slide_xml = base_slide_xml_str
        rels_xml = base_rels_xml_str

        site_label = f"Site {site_num}"
        # Shape 0: Clear the redundant small text above the badge so it is not repeated
        slide_xml = re.sub(r'(<a:t[^>]*>)Site\s*\d+(</a:t>)', r'\g<1>\g<2>', slide_xml, count=1)

        # Yellow badge (Shape 2): Highlight "Site 1", "Site 2" without any artificial site code
        slide_xml = slide_xml.replace('MB - 07', saxutils.escape(site_label))
        rels_xml = re.sub(r'Target="https://panel\.mediabuzzoutdoor\.com/campaigns\?site=[^"]*"',
                          'Target="https://panel.mediabuzzoutdoor.com/campaigns"',
                          rels_xml)

        city = entry.get('city') or 'Ahmedabad'
        slide_xml = slide_xml.replace('Ahmedabad', saxutils.escape(city))

        loc = entry.get('location') or 'Prime Billboard Location'
        slide_xml = slide_xml.replace('Shivranjani Bridge- Fcg. Shyamal 4 road Jnc - Left (1)', saxutils.escape(loc))

        size = entry.get('size') or '30x10 ft'
        mtype = entry.get('type') or 'Hoarding'
        illum = entry.get('illum') or 'BL'
        rate = entry.get('rate') or 'On Request'
        avail = entry.get('availability') or 'Available'
        coords = entry.get('coords') or '23.012345, 72.567890'

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

    pres_xml_str = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', f'<p:sldIdLst>{"".join(sld_id_list)}</p:sldIdLst>', pres_xml_str)

    files['[Content_Types].xml'] = content_types_str.encode('utf-8')
    files['ppt/presentation.xml'] = pres_xml_str.encode('utf-8')
    files['ppt/_rels/presentation.xml.rels'] = pres_rels_str.encode('utf-8')

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
        for fname, data in files.items():
            z_out.writestr(fname, data)


def process_presentation_pure(input_path, output_path, template_file):
    """Processes conversion using pure Python standard library only."""
    entries = extract_entries_pure_zip(input_path)
    if not entries:
        raise ValueError("Could not extract any location or billboard slides from the uploaded presentation.")
        
    generate_converted_pptx_pure(template_file, output_path, entries)
    
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
        'processedSlides': len(entries),
        'totalSlidesInFile': len(entries) + 3, # Cover 1, Cover 2, sites..., Closing
        'templateUsed': os.path.basename(template_file),
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

def process_presentation(input_path, output_path, master_template_path=None):
    """
    Main conversion entrypoint.
    Executes zero-dependency OpenXML engine ensuring 100% valid PowerPoint files.
    """
    template_file = find_master_template(master_template_path)
    if not template_file:
        raise FileNotFoundError("Media Buzz master template not found. Searched candidate paths.")
        
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input PowerPoint presentation not found at: {input_path}")
        
    return process_presentation_pure(input_path, output_path, template_file)


def main():
    parser = argparse.ArgumentParser(description="Media Buzz PowerPoint Converter Engine")
    parser.add_argument('--check-template', action='store_true', help="Check master template status")
    parser.add_argument('--input', type=str, help="Path to input presentation")
    parser.add_argument('--output', type=str, help="Path to output presentation")
    parser.add_argument('--template', type=str, help="Path to master template")
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
        
    if not args.input or not args.output:
        err_res = {'success': False, 'error': 'Both --input and --output are required for conversion.'}
        print(json.dumps(err_res, ensure_ascii=True))
        sys.exit(1)
        
    try:
        res = process_presentation(args.input, args.output, args.template)
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

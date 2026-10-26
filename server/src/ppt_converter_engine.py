#!/usr/bin/env python3
"""
Media Buzz OOH Workspace - PowerPoint Converter Engine
Converts any uploaded presentation into the official Media Buzz Master Template format
used in Automated PPT (MediaBuzz_Master_Template.pptx / Testing 24-10-2026.pptx).
"""

import os
import sys
import json
import re
import copy
import io
import base64
import argparse
from PIL import Image, ImageDraw, ImageFont

# Ensure utf-8 encoding for stdout
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

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
    """
    Extracts site code from text like MB-01, SSV-07, H-1111, ATH-001, SITE-03, LOC-12.
    Formats it cleanly like 'MB - 01' matching Automated PPT display standard.
    """
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
    """
    Extracts dimensions like 40x20 ft, 20x15, 61 x 20 FT, 30*10 ft.
    """
    m = re.search(r'\b(\d+(?:\.\d+)?\s*[\*xX×]\s*\d+(?:\.\d+)?(?:\s*(?:ft|feet|mtr|meter|m))?)\b', text, re.IGNORECASE)
    if m:
        val = m.group(1).strip()
        val = re.sub(r'\s*[\*X×]\s*', 'x', val)
        val = re.sub(r'\s*x\s*', 'x', val)
        if not re.search(r'(ft|feet|mtr|meter|m)$', val, re.IGNORECASE):
            val += ' ft'
        return val
    return ''

def clean_location(text, size_str='', code_str='', city_str=''):
    """
    Strips out site codes, size substrings, specification rows, and trailing illumination.
    """
    s = text.strip()
    
    # Cut off before specification labels (Size, Type, Illumination, Rate, etc.)
    s = re.split(r'\b(?:Size|Type|Illumination|Adv\.?fee|Rate|Availability|Location Coordinates|Coordinates)\s*:', s, flags=re.IGNORECASE)[0].strip()
    
    # Strip common site code prefixes: ATH-001, H-1111, MB-09, SSV-05, LOC-12, SITE 1, etc.
    s = re.sub(r'^[A-Z0-9]{1,7}\s*[-–—]\s*\d+\s*[-–—)]*\s*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^(?:Site|Location|Loc|Sr\.?\s*No\.?)\s*#?\d+\s*[-–—:]*\s*', '', s, flags=re.IGNORECASE)
    
    # Strip code_str if present
    if code_str:
        raw_code = code_str.replace(' - ', '-').strip()
        s = re.sub(rf'\b{re.escape(raw_code)}\b', '', s, flags=re.IGNORECASE)
        s = re.sub(rf'\b{re.escape(code_str)}\b', '', s, flags=re.IGNORECASE)
        
    # Strip city prefix if present
    if city_str:
        s = re.sub(rf'^{re.escape(city_str)}\s*[-–—:,]*\s*', '', s, flags=re.IGNORECASE)

    # Strip size substring if present in the text
    if size_str:
        raw_size = re.sub(r'\s*(?:ft|feet|mtr|meter|m)$', '', size_str, flags=re.IGNORECASE).strip()
        raw_pat = re.sub(r'x', r'[*xX×]', raw_size)
        s = re.sub(rf'[-–—(\[]*\s*{raw_pat}(?:\s*(?:ft|feet|mtr|meter|m))?\s*[-–—)\]]*', '', s, flags=re.IGNORECASE)
        s = re.sub(rf'[-–—(\[]*\s*{re.escape(size_str)}\s*[-–—)\]]*', '', s, flags=re.IGNORECASE)
        
    # Strip trailing illumination codes: FL, BL, NL, Front Lit, Back Lit, Non Lit
    s = re.sub(r'[\s,\-–—]+(FL|BL|NL|Front\s*Lit|Back\s*Lit|Non\s*Lit)\s*$', '', s, flags=re.IGNORECASE)
    
    # Clean leading/trailing punctuation and repeated spaces
    s = re.sub(r'^[\s\-–—:.)\]]+', '', s)
    s = re.sub(r'[\s\-–—:.(]+$', '', s)
    s = re.sub(r'\s{2,}', ' ', s)
    return s.strip()

def detect_city(text):
    match = re.search(r'\b(Ahmedabad|Surat|Vadodara|Rajkot|Gandhinagar|Bhavnagar|Jamnagar|Mumbai|Pune|Delhi|Bengaluru|Bangalore|Hyderabad|Chennai|Kolkata)\b', text, re.IGNORECASE)
    if match:
        return match.group(1).capitalize()
    return 'Ahmedabad'

def detect_media_type(text):
    m = re.search(r'\b(Gantry|Unipole|Hoarding|Billboard|Foot\s*Over\s*Bridge|FOB|Bus\s*Shelter|Cantilever|Pole\s*Kiosk|Kiosk)\b', text, re.IGNORECASE)
    if m:
        val = m.group(1).title()
        if val.upper() in ['FOB', 'FOOT OVER BRIDGE']:
            return 'FOB'
        return val
    return 'Hoarding'

def detect_illumination(text):
    if re.search(r'\b(Front\s*Lit|FL)\b', text, re.IGNORECASE):
        return 'FL'
    if re.search(r'\b(Back\s*Lit|BL)\b', text, re.IGNORECASE):
        return 'BL'
    if re.search(r'\b(Non\s*Lit|NL)\b', text, re.IGNORECASE):
        return 'Non Lit'
    if re.search(r'\b(Lit)\b', text, re.IGNORECASE):
        return 'LIT'
    return 'BL'

def detect_rate(text):
    m = re.search(r'(?:₹|Rs\.?|INR)\s*([\d,]+(?:\.\d+)?(?:\s*/-)?(?:\s*(?:pm|per\s*month))?)', text, re.IGNORECASE)
    if m:
        num = re.sub(r'[^\d,]', '', m.group(1)).strip()
        if num:
            return f"₹ {num}/-"
    return 'On Request'

def detect_availability(text):
    m = re.search(r'\b(Booked\s*till\s*[\w\s\d,.-]+|Available\s*from\s*[\w\s\d,.-]+|Available|Occupied|Vacant)\b', text, re.IGNORECASE)
    if m:
        val = m.group(1).strip()
        if re.search(r'^available\s+from\s+', val, re.IGNORECASE):
            val = re.sub(r'^available\s+from\s+', 'Booked till ', val, flags=re.IGNORECASE)
        return val
    return 'Available'

def detect_coords(text):
    m = re.search(r'\b(2[0-4]\.\d{4,8}\s*,\s*7[0-4]\.\d{4,8})\b', text)
    if m:
        return re.sub(r'\s+', ' ', m.group(1).strip())
    m2 = re.search(r'\b(\d{1,2}\.\d{4,8}\s*,\s*\d{1,2}\.\d{4,8})\b', text)
    if m2:
        return re.sub(r'\s+', ' ', m2.group(1).strip())
    return '23.012345, 72.567890'

def generate_no_photo_bytes():
    """Generates a clean 1500x1224 dark card image for missing photos matching Automated PPT."""
    w, h = 1500, 1224
    img = Image.new('RGB', (w, h), color=(17, 17, 17))
    draw = ImageDraw.Draw(img)
    # Border
    draw.rectangle([10, 10, w - 10, h - 10], outline=(42, 42, 42), width=3)
    text = "NO SITE PHOTO UPLOADED"
    # Centered text
    draw.text((w // 2 - 250, h // 2 - 20), text, fill=(85, 85, 85))
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    return buf.getvalue()

def duplicate_slide(prs, src_slide):
    """
    Cleanly duplicate a slide within the presentation preserving all layouts, shapes, and relationships.
    """
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
    Generates a visual 16:9 JPEG preview matching the exact Automated PPT master theme.
    """
    w, h = 800, 450
    canvas = Image.new('RGB', (w, h), color=(0, 0, 0))
    draw = ImageDraw.Draw(canvas)
    
    # Left: Photo card at px=42, py=42, pw=375, ph=306
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
        
    # Official Media Buzz logo badge at bottom-left
    lb_x, lb_y, lb_w, lb_h = 42, 363, 111, 45
    draw.rounded_rectangle([lb_x, lb_y, lb_x + lb_w, lb_y + lb_h], radius=8, fill=(255, 194, 0))
    draw.text((lb_x + 14, lb_y + 8), "MEDIA BUZZ", fill=(3, 3, 82))
    draw.text((lb_x + 28, lb_y + 24), "Be Seen", fill=(3, 3, 82))
    
    # Right: Information panel
    rx = 447
    draw.text((rx, 42), f"Site {site_num}", fill=(255, 255, 255))
    
    # Yellow badge for Site Code (e.g. MB - 01)
    bx, by, bw, bh = rx, 63, 160, 43
    draw.rounded_rectangle([bx, by, bx+bw, by+bh], radius=8, fill=(255, 194, 0))
    draw.text((bx + 18, by + 12), code or f"MB - {site_num:02d}", fill=(0, 0, 0))
    
    # City (Bold pure white)
    draw.text((rx, 118), city or "Ahmedabad", fill=(255, 255, 255))
        
    # Location description
    draw.text((rx, 148), location[:44] if location else "Location Details", fill=(255, 255, 255))
    if len(location) > 44:
        draw.text((rx, 168), location[44:88], fill=(255, 255, 255))
        
    # 6 Specification rows matching Automated PPT
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


def process_presentation(input_path, output_path, master_template_path=None):
    if not pptx:
        raise RuntimeError("python-pptx library is not installed.")
        
    template_file = find_master_template(master_template_path)
    if not template_file:
        raise FileNotFoundError("Media Buzz master template not found. Searched candidate paths.")
        
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input PowerPoint presentation not found at: {input_path}")
        
    in_prs = pptx.Presentation(input_path)
    out_prs = pptx.Presentation(template_file)
    
    # Ensure template has clean structure: Cover 1, Cover 2, Site Slide Template, Closing Slide
    # If template has 5 slides (due to extra legacy demo site slide at index 3), remove it cleanly
    if len(out_prs.slides) == 5:
        demo_sldId = out_prs.slides._sldIdLst[3]
        out_prs.part.drop_rel(demo_sldId.rId)
        out_prs.slides._sldIdLst.remove(demo_sldId)

    if len(out_prs.slides) < 3:
        raise ValueError("Master template does not contain the required site slide at slide index 3.")
        
    base_site_slide = out_prs.slides[2]
    closing_sldId = out_prs.slides._sldIdLst[len(out_prs.slides) - 1]
    
    # 1. Extract location entries from input presentation
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
        
        # Skip pure closing / contact / thank you slides
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

        # Find best billboard photograph
        best_pic = None
        max_area = 0
        for sh in slide.shapes:
            if sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
                w_in = sh.width / 914400
                h_in = sh.height / 914400
                area = w_in * h_in
                
                # Filter out full-bleed slide background images (covers)
                if w_in >= 0.92 * slide_w_in and h_in >= 0.92 * slide_h_in:
                    continue
                # Filter out small icons/badges
                if area > max_area and area > 2.0:
                    max_area = area
                    best_pic = sh
                    
        # If slide has no photo, no size, and no location, skip it (cover / intro slides)
        if not best_pic and not size and (not loc or len(loc) < 5):
            continue
        
        # Review flags
        review_flags = []
        if not best_pic:
            review_flags.append("Photo missing or not detected")
        if not size:
            review_flags.append("Dimensions/size missing")
        if not loc or len(loc) < 3:
            review_flags.append("Location description unclear")
            
        entries.append({
            'source_slide': i + 1,
            'photo_bytes': best_pic.image.blob if best_pic else None,
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
    
    # 2. Populate slide matching Automated PPT theme
    def apply_data_to_slide(slide, entry, site_num):
        # Update or set Photo (Image 0 or main billboard picture)
        photo_data = entry.get('photo_bytes') or generate_no_photo_bytes()
        try:
            buf = io.BytesIO(photo_data)
            new_img_part, rId = slide.part.get_or_add_image_part(buf)
            pic_sh = None
            for sh in slide.shapes:
                if sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE:
                    w_in = sh.width / 914400
                    l_in = sh.left / 914400
                    if w_in > 4.0 and l_in < 2.0:
                        pic_sh = sh
                        break
            if not pic_sh:
                pic_sh = slide.shapes[0]
            blip = pic_sh._element.xpath('.//a:blip')[0]
            blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', rId)
        except Exception as e:
            entry['review_needed'] = True
            entry['review_flags'].append(f"Photo insertion notice: {str(e)}")

        code_text = entry.get('code') or f"MB - {site_num:02d}"
        size_text = entry.get('size') or '30x10 ft'
        city_text = entry.get('city') or 'Ahmedabad'
        loc_text = entry.get('location') or 'Location Details'
        type_text = entry.get('type') or 'Hoarding'
        illum_text = entry.get('illum') or 'BL'
        rate_text = entry.get('rate') or 'On Request'
        avail_text = entry.get('availability') or 'Available'
        coords_text = entry.get('coords') or '23.012345, 72.567890'

        def set_spec_paragraph(p, label, val):
            if len(p.runs) >= 2:
                p.runs[0].text = f"{label}:  "
                p.runs[0].font.name = 'Arial'
                p.runs[0].font.size = pptx.util.Pt(14.5)
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = pptx.dml.color.RGBColor(255, 194, 0) # Gold
                
                p.runs[1].text = val
                p.runs[1].font.name = 'Arial'
                p.runs[1].font.size = pptx.util.Pt(14.5)
                p.runs[1].font.bold = True
                p.runs[1].font.color.rgb = pptx.dml.color.RGBColor(255, 255, 255) # Pure white
            elif len(p.runs) == 1:
                p.runs[0].text = f"{label}:  "
                p.runs[0].font.name = 'Arial'
                p.runs[0].font.size = pptx.util.Pt(14.5)
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = pptx.dml.color.RGBColor(255, 194, 0)
                
                r2 = p.add_run()
                r2.text = val
                r2.font.name = 'Arial'
                r2.font.size = pptx.util.Pt(14.5)
                r2.font.bold = True
                r2.font.color.rgb = pptx.dml.color.RGBColor(255, 255, 255)

        # Update texts preserving formatting runs
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            name = sh.name
            tf = sh.text_frame
            top_in = sh.top / 914400
            left_in = sh.left / 914400
            w_in = sh.width / 914400

            # 1. Site Counter ("Site 1")
            if name == 'Text 0' or (abs(top_in - 0.70) < 0.15 and abs(left_in - 7.45) < 0.25):
                if tf.paragraphs and tf.paragraphs[0].runs:
                    p = tf.paragraphs[0]
                    p.runs[0].text = f"Site {site_num}"
                    p.runs[0].font.name = 'Arial'
                    p.runs[0].font.size = pptx.util.Pt(13)
                    p.runs[0].font.color.rgb = pptx.dml.color.RGBColor(255, 255, 255)
                else:
                    tf.text = f"Site {site_num}"
                    
            # 2. Yellow Pill Badge: Site Code ("MB - 01")
            elif name == 'Text 2' or (abs(top_in - 1.05) < 0.15 and abs(left_in - 7.45) < 0.25 and 'text' in name.lower()):
                p = tf.paragraphs[0]
                p.text = code_text
                if p.runs:
                    p.runs[0].font.name = 'Arial'
                    p.runs[0].font.size = pptx.util.Pt(26)
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = pptx.dml.color.RGBColor(0, 0, 0) # Bold Black
                    
            # 3. City Heading ("Ahmedabad")
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
                    
            # 4. Location / Address Subtitle
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

            # 5. Specification Rows
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

    # 3. Populate site slides (Slide index 2 is base site slide)
    for idx in range(n_entries):
        site_num = idx + 1
        if idx == 0:
            slide = out_prs.slides[2]
        else:
            slide = duplicate_slide(out_prs, base_site_slide)
        apply_data_to_slide(slide, entries[idx], site_num)

    # Always ensure the closing slide is positioned at the very end of the presentation
    out_prs.slides._sldIdLst.remove(closing_sldId)
    out_prs.slides._sldIdLst.append(closing_sldId)
        
    # Save output presentation
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    out_prs.save(output_path)
    
    # 4. Generate previews and output data matching Automated PPT theme
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
        print(json.dumps(res))
        return
        
    if not args.input or not args.output:
        print(json.dumps({'success': False, 'error': 'Both --input and --output are required for conversion.'}))
        sys.exit(1)
        
    try:
        res = process_presentation(args.input, args.output, args.template)
        print(json.dumps(res))
    except Exception as e:
        err_res = {
            'success': False,
            'error': str(e)
        }
        print(json.dumps(err_res))
        sys.exit(1)

if __name__ == '__main__':
    main()

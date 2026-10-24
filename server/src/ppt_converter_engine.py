#!/usr/bin/env python3
"""
PowerPoint Converter Engine for Media Buzz OOH Workspace
Converts any uploaded presentation into the master Testing PPT format (Testing 24-10-2026.pptx).
"""

import os
import sys
import json
import re
import copy
import io
import base64
import argparse
from PIL import Image, ImageDraw

# Ensure utf-8 encoding for stdout
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

try:
    import pptx
except ImportError:
    pptx = None


CANDIDATE_TEMPLATE_PATHS = [
    os.path.join(os.path.dirname(__file__), '..', 'templates', 'Testing 24-10-2026.pptx'),
    os.path.join(os.path.dirname(__file__), '..', 'templates', 'Testing  24-10-2026.pptx'),
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

def extract_size(text):
    """
    Extracts dimensions like 40x20 ft, 20x15, 61 x 20 FT, 30*10 ft.
    """
    m = re.search(r'\b(\d+(?:\.\d+)?\s*[\*xX×]\s*\d+(?:\.\d+)?(?:\s*(?:ft|feet|mtr|meter|m))?)\b', text, re.IGNORECASE)
    if m:
        val = m.group(1).strip()
        # Normalize multiplication symbol
        val = re.sub(r'\s*[\*X×]\s*', 'x', val)
        val = re.sub(r'\s*x\s*', 'x', val)
        # Ensure 'ft' suffix if no unit
        if not re.search(r'(ft|feet|mtr|meter|m)$', val, re.IGNORECASE):
            val += ' ft'
        return val
    return ''

def clean_location(text, size_str=''):
    """
    Strips out site codes, inventory IDs, size substrings, and trailing illumination.
    """
    s = text.strip()
    
    # Strip common site code prefixes: ATH-001, H-1111, MB-09, SSV-05, LOC-12, SITE 1, etc.
    s = re.sub(r'^[A-Z0-9]{1,7}\s*[-–—]\s*\d+\s*[-–—)]*\s*', '', s, flags=re.IGNORECASE)
    s = re.sub(r'^(?:Site|Location|Loc|Sr\.?\s*No\.?)\s*#?\d+\s*[-–—:]*\s*', '', s, flags=re.IGNORECASE)
    
    # Strip size substring if present in the text
    if size_str:
        raw_size = re.sub(r'\s*(?:ft|feet|mtr|meter|m)$', '', size_str, flags=re.IGNORECASE).strip()
        # e.g. 40x20, 40*20
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
    return ''

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

def generate_slide_thumbnail(photo_bytes, site_num, size, location, city=""):
    """
    Generates a visual 16:9 JPEG preview matching the Testing PPT layout.
    """
    w, h = 800, 450
    canvas = Image.new('RGB', (w, h), color=(10, 15, 22))
    draw = ImageDraw.Draw(canvas)
    
    # Left: Photo placeholder/content
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
        draw.text((px + 90, py + 140), "No Photo Available", fill=(120, 130, 145))
        
    # Brand logo area at bottom-left
    draw.text((42, 370), "MEDIA BUZZ OUTDOOR", fill=(255, 194, 0))
    
    # Right: Information panel
    rx = 447
    draw.text((rx, 42), f"Site {site_num}", fill=(255, 255, 255))
    
    # Yellow badge for size
    bx, by, bw, bh = rx, 63, 160, 43
    draw.rounded_rectangle([bx, by, bx+bw, by+bh], radius=8, fill=(255, 194, 0))
    draw.text((bx + 14, by + 12), size or "SIZE", fill=(0, 0, 0))
    
    if city:
        draw.text((rx, 118), city, fill=(255, 255, 255))
        
    # Location
    draw.text((rx, 148), location[:42], fill=(255, 255, 255))
    if len(location) > 42:
        draw.text((rx, 168), location[42:84], fill=(255, 255, 255))
        
    # Fields (preserved labels, blank values)
    fields = [
        ("Size:", size),
        ("Type:", ""),
        ("Illumination:", ""),
        ("Adv.fee per month:", ""),
        ("Availability:", ""),
        ("Location Coordinates:", "")
    ]
    
    fy = 200
    for label, val in fields:
        draw.text((rx, fy), label, fill=(255, 194, 0))
        if val:
            draw.text((rx + 110, fy), val, fill=(255, 255, 255))
        draw.line([rx, fy + 20, rx + 310, fy + 20], fill=(35, 45, 55), width=1)
        fy += 32
        
    buf = io.BytesIO()
    canvas.save(buf, format='JPEG', quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode('utf-8')


def process_presentation(input_path, output_path, master_template_path=None):
    if not pptx:
        raise RuntimeError("python-pptx library is not installed.")
        
    template_file = find_master_template(master_template_path)
    if not template_file:
        raise FileNotFoundError(f"Master template 'Testing 24-10-2026.pptx' not found. Searched candidate paths.")
        
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input PowerPoint presentation not found at: {input_path}")
        
    in_prs = pptx.Presentation(input_path)
    out_prs = pptx.Presentation(template_file)
    
    # Slide 3 (index 2) is the site layout in Testing 24-10-2026.pptx
    if len(out_prs.slides) < 3:
        raise ValueError("Master template does not contain the required site slide at slide index 3.")
    base_site_slide = out_prs.slides[2]
    
    # Closing slide element in template (last slide)
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
        
        # Skip pure closing/thank you slides
        if full_text.lower() in ['mediabuzzoutdoor.com', 'thank you', 'thanks', 'contact us']:
            continue
            
        size = extract_size(full_text)
        loc = clean_location(full_text, size)
        city = detect_city(full_text)

        # Find best photo
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
                    
        # If slide has no photo, no size, and no location (e.g. cover/intro slide), skip it
        if not best_pic and not size and (not loc or len(loc) < 5):
            continue
        
        # Check review flags
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
            'size': size,
            'location': loc,
            'city': city,
            'review_needed': len(review_flags) > 0,
            'review_flags': review_flags
        })
        
    if not entries:
        raise ValueError("Could not extract any location or billboard slides from the uploaded presentation.")
        
    n_entries = len(entries)
    
    # 2. Populate slides
    def apply_data_to_slide(slide, entry, site_num):
        # Update Photo (Image 0 or main billboard picture)
        if entry['photo_bytes']:
            try:
                buf = io.BytesIO(entry['photo_bytes'])
                new_img_part, rId = slide.part.get_or_add_image_part(buf)
                pic_sh = None
                for sh in slide.shapes:
                    if sh.name == 'Image 0' or (sh.shape_type == pptx.enum.shapes.MSO_SHAPE_TYPE.PICTURE and (sh.width / 914400) > 4.0):
                        pic_sh = sh
                        break
                if not pic_sh:
                    pic_sh = slide.shapes[0]
                blip = pic_sh._element.xpath('.//a:blip')[0]
                blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', rId)
            except Exception as e:
                entry['review_needed'] = True
                entry['review_flags'].append(f"Photo insertion failed: {str(e)}")
                
        # Update texts preserving formatting runs
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            name = sh.name
            tf = sh.text_frame
            
            if name == 'Text 0': # Site title
                if tf.paragraphs and tf.paragraphs[0].runs:
                    tf.paragraphs[0].runs[0].text = f"Site {site_num}"
                else:
                    tf.text = f"Site {site_num}"
                    
            elif name == 'Text 2': # Rounded yellow box heading
                if tf.paragraphs and tf.paragraphs[0].runs:
                    tf.paragraphs[0].runs[0].text = entry['size'] or ""
                else:
                    tf.text = entry['size'] or ""
                    
            elif name == 'Text 3': # City / Area
                if tf.paragraphs and tf.paragraphs[0].runs:
                    tf.paragraphs[0].runs[0].text = entry['city'] or ""
                else:
                    tf.text = entry['city'] or ""
                    
            elif name == 'Text 4': # Location description
                if tf.paragraphs and tf.paragraphs[0].runs:
                    tf.paragraphs[0].runs[0].text = entry['location'] or ""
                else:
                    tf.text = entry['location'] or ""
                    
            elif name == 'Text 5': # Size row
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    if len(p.runs) >= 2:
                        p.runs[0].text = "Size:  "
                        p.runs[1].text = entry['size'] or ""
                    elif len(p.runs) == 1:
                        p.runs[0].text = f"Size:  {entry['size'] or ''}"
                        
            elif name == 'Text 7': # Type row -> preserve label, blank value
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    if len(p.runs) >= 2:
                        p.runs[0].text = "Type:  "
                        p.runs[1].text = ""
                    elif len(p.runs) == 1:
                        p.runs[0].text = "Type:  "
                        
            elif name == 'Text 9': # Illumination row -> preserve label, blank value
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    if len(p.runs) >= 2:
                        p.runs[0].text = "Illumination:  "
                        p.runs[1].text = ""
                    elif len(p.runs) == 1:
                        p.runs[0].text = "Illumination:  "
                        
            elif name == 'Text 11': # Adv fee row -> preserve label, blank value
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    if len(p.runs) >= 2:
                        p.runs[0].text = "Adv.fee per month:  "
                        p.runs[1].text = ""
                    elif len(p.runs) == 1:
                        p.runs[0].text = "Adv.fee per month:  "
                        
            elif name == 'Text 13': # Availability row -> preserve label, blank value
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    if len(p.runs) >= 2:
                        p.runs[0].text = "Availability:  "
                        p.runs[1].text = ""
                    elif len(p.runs) == 1:
                        p.runs[0].text = "Availability:  "
                        
            elif name == 'Text 15': # Coordinates row -> preserve label, blank value
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    if len(p.runs) >= 2:
                        p.runs[0].text = "Location Coordinates:  "
                        p.runs[1].text = ""
                    elif len(p.runs) == 1:
                        p.runs[0].text = "Location Coordinates:  "

    # Identify site slides in template (between cover slides and closing slide)
    total_tmpl_slides = len(out_prs.slides)
    existing_site_slides = max(1, total_tmpl_slides - 3)
    
    # Populate existing site slides
    populate_count = min(n_entries, existing_site_slides)
    for idx in range(populate_count):
        slide = out_prs.slides[2 + idx]
        apply_data_to_slide(slide, entries[idx], idx + 1)
        
    # If fewer entries than existing site slides, cleanly delete excess template site slides
    if n_entries < existing_site_slides:
        for idx in range(2 + existing_site_slides - 1, 1 + n_entries, -1):
            sldId = out_prs.slides._sldIdLst[idx]
            out_prs.part.drop_rel(sldId.rId)
            out_prs.slides._sldIdLst.remove(sldId)
            
    # If more entries than existing site slides, duplicate base slide and update
    elif n_entries > existing_site_slides:
        for idx in range(existing_site_slides, n_entries):
            new_slide = duplicate_slide(out_prs, base_site_slide)
            apply_data_to_slide(new_slide, entries[idx], idx + 1)
            
        # Re-position closing slide to the end of presentation
        out_prs.slides._sldIdLst.remove(closing_sldId)
        out_prs.slides._sldIdLst.append(closing_sldId)
        
    # Save output presentation
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    out_prs.save(output_path)
    
    # 3. Generate previews and output data
    slide_previews = []
    for idx, entry in enumerate(entries):
        site_num = idx + 1
        thumb_data = generate_slide_thumbnail(
            entry['photo_bytes'],
            site_num,
            entry['size'],
            entry['location'],
            entry['city']
        )
        slide_previews.append({
            'siteNumber': site_num,
            'sourceSlide': entry['source_slide'],
            'size': entry['size'],
            'location': entry['location'],
            'city': entry['city'],
            'hasPhoto': entry['photo_bytes'] is not None,
            'reviewNeeded': entry['review_needed'],
            'reviewFlags': entry['review_flags'],
            'previewUrl': thumb_data
        })
        
    return {
        'success': True,
        'message': f'Successfully converted presentation to Testing format.',
        'processedSlides': n_entries,
        'totalSlidesInFile': len(out_prs.slides),
        'templateUsed': os.path.basename(template_file),
        'slides': slide_previews
    }


def main():
    parser = argparse.ArgumentParser(description="PowerPoint Converter Engine")
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

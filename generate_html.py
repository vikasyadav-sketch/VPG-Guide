import docx
from jinja2 import Template
import re
import os
from docx.oxml.ns import qn
from datetime import datetime

def find_car_images(images_folder=None):
    """
    Find car view images (Front, Side, Rear, Quarter) in the car images folder.
    Returns a dict with keys: 'front', 'side', 'rear', 'quarter'
    Each value is the full URL to the image or empty string if not found.
    Image filename pattern: Car images contain view keywords.
    Examples: MBAMGGLA352.0LFron.jpeg, xyzfront.jpg, xyz-front.png
    """
    car_images = {
        'front': '',
        'side': '',
        'rear': '',
        'quarter': ''
    }
    
    # Try to find the car images folder with various naming conventions
    possible_folder_names = [
        'Car images',
        'Car image', 
        'Car',
        'car images',
        'car image',
        'car',
        'Images',
        'images',
        'Image',
        'image'
    ]
    
    folder_to_use = None
    
    # If a specific folder is provided, use it
    if images_folder and os.path.exists(images_folder):
        folder_to_use = images_folder
    else:
        # Search for the folder with various names
        for folder_name in possible_folder_names:
            if os.path.exists(folder_name):
                folder_to_use = folder_name
                break
    
    if not folder_to_use:
        return car_images
    
    # Get all image files
    image_files = [f for f in os.listdir(folder_to_use) 
                   if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
    
    if not image_files:
        return car_images
    
    # Define view patterns with variations for fuzzy matching
    view_patterns = {
        'front': ['front', 'fron', 'fro', 'frnt'],
        'side': ['side', 'sid'],
        'rear': ['rear', 'rea'],
        'quarter': ['quarter', 'quattr', 'quater', 'quatr', 'quar' ,'qua','quat']
    }
    
    # Search for car view images
    for img_file in image_files:
        filename_lower = img_file.lower()
        # Remove extension and normalize (remove hyphens, underscores, spaces)
        filename_without_ext = os.path.splitext(filename_lower)[0]
        filename_normalized = filename_without_ext.replace('-', '').replace('_', '').replace(' ', '')
        
        # Check each view type
        for view_type, patterns in view_patterns.items():
            # Skip if we already found this view
            if car_images[view_type]:
                continue
            
            # Check if any pattern matches
            for pattern in patterns:
                # Check if pattern appears at the end or anywhere in the filename
                if (filename_normalized.endswith(pattern) or 
                    pattern in filename_normalized):
                    car_images[view_type] = f'https://admin.eeuroparts.com/var/theme/images/{img_file}'
                    break
    
    return car_images

def clean_url(url):
    """
    Clean URL by removing query parameters and fragments.
    Returns the base URL without tracking parameters.
    """
    # Remove query parameters (everything after ?)
    if '?' in url:
        url = url.split('?')[0]
    
    # Remove fragments (everything after #)
    if '#' in url:
        url = url.split('#')[0]
    
    return url.strip()

def extract_hyperlinks_from_paragraph(paragraph, doc):
    """
    Extract all hyperlinks from a paragraph.
    Returns a list of tuples: [(hyperlink_text, url), ...]
    URLs are cleaned to remove query parameters and fragments.
    """
    hyperlinks = []
    
    # Find all hyperlink elements in the paragraph
    hyperlink_elements = paragraph._element.findall('.//' + qn('w:hyperlink'))
    
    for hl_elem in hyperlink_elements:
        # Get the text from the hyperlink
        texts = [t.text for t in hl_elem.findall('.//' + qn('w:t')) if t.text]
        hyperlink_text = ''.join(texts) if texts else ''
        
        # Get the relationship ID
        r_id = hl_elem.get(qn('r:id'))
        
        if r_id and hyperlink_text:
            try:
                # Get the actual URL from the relationship
                url = doc.part.rels[r_id].target_ref
                # Clean the URL to remove any query parameters or fragments
                clean_url_str = clean_url(url)
                hyperlinks.append((hyperlink_text.strip(), clean_url_str))
            except:
                pass
    
    return hyperlinks

def parse_word_document(docx_path):
    doc = docx.Document(docx_path)
    # Replace en-dashes and em-dashes with normal hyphens
    # IMPORTANT: Split by \n to handle multi-field paragraphs
    # Also store paragraph objects for hyperlink extraction
    paragraphs = []
    paragraph_objects = []
    for p in doc.paragraphs:
        text = p.text.strip().replace('–', '-').replace('—', '-')
        if text:
            # Split by \n to handle cases where multiple fields are in one paragraph
            for line in text.split('\n'):
                line = line.strip()
                if line:
                    paragraphs.append(line)
                    paragraph_objects.append(p)  # Store the paragraph object for hyperlink extraction

    data = {
        'vehicle_heading': '',
        'description_text': '',
        'common_issues_heading': '',
        'car_images': {},
        'specs': {},
        'issues': {
            'Brakes': [],
            'Suspension': [],
            'Ignition': [],
            'Steering': [],
            'Engine': [],
            'Fuel Delivery': [],
            'Electrical System': [],
            'Driveline/Transmission': [],
            'Others': []
        }
    }

    # Find car images
    data['car_images'] = find_car_images()

# 1. Extract FULL Heading (do NOT trim)
vpg_index = -1
for i, p in enumerate(paragraphs):
    if p.lower().startswith('vehicle platform guide'):
        data['vehicle_heading'] = p.strip()
        vpg_index = i
        break
# Fallback
if vpg_index == -1 and paragraphs:
    data['vehicle_heading'] = paragraphs[0].strip()
    vpg_index = 0
    
# 2. Extract FULL Description (multiple paragraphs)
description_paragraphs = []

if vpg_index != -1:
    for i in range(vpg_index + 1, len(paragraphs)):
        p = paragraphs[i].strip()

        # Stop when specs or structured sections start
        if any(x in p for x in [
            'Specifications',
            'Common Issues',
            'Fault Codes',
            'Top 20'
        ]):
            break

        if len(p) > 40:
            description_paragraphs.append(p)

data['description_text'] = '\n\n'.join(description_paragraphs)

    # Extract Specifications (Preserving existing logic)
    data['specs'] = {
        'Engine and Powertrain': {},
        'Fuel Economy (EPA Estimates)': {},
        'Vehicle Weight': {},
        'Configurations and Submodels': {},
        # 'Other Specifications': {} # Removed as per user request
    }
# /* ========================= GALLERY (BASE) ========================= */
    def categorize_spec(key, value):
        key_lower = key.lower()
 
        if any(x in key_lower for x in ['engine', 'horse', 'torque', 'transmission', 'fuel type', 'displacement', 'cylinders']):
            return 'Engine and Powertrain'
        elif any(x in key_lower for x in ['drive', 'configuration', 'submodel', 'trim', 'body', 'door', 'seat' , 'capacity']):
            return 'Configurations and Submodels'
        elif any(x in key_lower for x in ['mpg', 'fuel economy', 'city', 'highway', 'combined']):
            return 'Fuel Economy (EPA Estimates)'
        elif any(x in key_lower for x in ['weight', 'payload', 'towing', 'gvwr']):
            return 'Vehicle Weight'
        return None # Return None for 'Other Specifications' to skip them

    # 3. Extract Heading before Category Issue
    # "Top Common Issues with the Audi Q5..."
    common_issues_index = -1
    for i, p in enumerate(paragraphs):
        if 'Top Common Issues' in p or 'Common Issues' in p:
            data['common_issues_heading'] = p
            common_issues_index = i
            break
    
    if common_issues_index == -1:
        common_issues_index = vpg_index + 5 # Fallback

    # Scan for specs
    # Extract Specifications (Preserving existing logic)
    # Limit scan to before common issues to avoid picking up symptoms as specs
    scan_limit = common_issues_index if common_issues_index != -1 else min(50, len(paragraphs))
    
    for j in range(0, scan_limit):
        p = paragraphs[j].strip()
        if ':' in p and len(p) < 200:
            if not any(x in p for x in ['Vehicle Platform Guide', 'In this', 'Common Issues', 'Fault Codes:', 'Why it happens:', 'Symptoms:', 'Parts to Replace:', 'Brands:']):
                parts = p.split(':', 1)
                if len(parts) == 2:
                    key = parts[0].strip()
                    val = parts[1].strip()
                    if key and val and len(key) < 50 and not key.startswith('Note'):
                        category = categorize_spec(key, val)
                        if category: # Only add if category is valid (not None)
                            data['specs'][category][key] = val

    # 4. Categories and Issues
    # Categories: Brakes, Suspension, Ignition, Steering, Engine, Fuel Delivery, Electrical System, Driveline/Transmission, Others
    # Naming convention: Category Name + System
    
    category_map = {
        'Brake System': 'Brakes',
        'Brakes System': 'Brakes',
        'Brakes': 'Brakes',
        'Suspension System': 'Suspension',
        'Suspension': 'Suspension',
        'Ignition System': 'Ignition',
        'Ignition': 'Ignition',
        'Steering System': 'Steering',
        'Steering': 'Steering',
        'Engine Management System': 'Engine',
        'Engine System': 'Engine',
        'Engine': 'Engine',
        'Fuel Delivery System': 'Fuel Delivery',
        'Fuel System': 'Fuel Delivery',
        'Fuel Delivery': 'Fuel Delivery',
        'Electrical Management System': 'Electrical System',
        'Electrical Systems': 'Electrical System',
        'Electrical System': 'Electrical System',
        'Electrical': 'Electrical System',
        'Driveline': 'Driveline/Transmission',
        'Driveline System': 'Driveline/Transmission',
        'Transmission System': 'Driveline/Transmission',
        'Driveline/Transmission System': 'Driveline/Transmission',
        'Transmission': 'Driveline/Transmission',
        'Driveline / Transmission': 'Driveline/Transmission',
        'Driveline / Transmission System': 'Driveline/Transmission',
        'Other System': 'Others',
        'Others': 'Others'
    }
    
    category_keys_lower = set(k.lower() for k in category_map.keys())

    current_category = None
    i = common_issues_index + 1
    
    while i < len(paragraphs):
        p = paragraphs[i].strip()
        
        # Check for Category Header
        is_category = False
        p_clean = p.strip().lower().rstrip(':')
        if p_clean in category_keys_lower:
            current_category = category_map.get(p.strip().rstrip(':'), None) # Try exact match first
            if not current_category:
                # Find the key that matches case-insensitively
                for k, v in category_map.items():
                    if k.lower() == p_clean:
                        current_category = v
                        break
            is_category = True
        
        if is_category:
            i += 1
            continue
            
        if current_category:
            # Parse Issue
            # Structure: Title -> Fault Codes -> Why -> Symptoms -> Parts -> Brands
            
            title_text = p
            fault_codes_inline = ''
            why_inline = ''
            symptoms_inline = []
            
            # Helper to extract fault codes
            def extract_fault_codes(text):
                return text

            # Check if multiple fields are on the title line
            # Split by field keywords
            remaining_text = title_text
            
            # Extract Fault Codes if present
            fc_match = re.search(r'(Fault Codes?|Fault Code)[\s:\-]+(.+?)(?=(Why it happens|Symptoms|Parts to Replace|Brands|$))', remaining_text, re.IGNORECASE)
            if fc_match:
                title_text = remaining_text[:fc_match.start()].strip()
                fault_codes_inline = fc_match.group(2).strip()
                remaining_text = remaining_text[fc_match.end():]
            
            # Extract Why it happens if present
            why_match = re.search(r'(Why it happens)[\s:\-]+(.+?)(?=(Symptoms|Parts to Replace|Brands|$))', remaining_text, re.IGNORECASE)
            if why_match:
                if not fc_match:  # Only update title if we haven't already
                    title_text = remaining_text[:why_match.start()].strip()
                why_inline = why_match.group(2).strip()
                remaining_text = remaining_text[why_match.end():]
            
            # Extract Symptoms if present on same line
            sym_match = re.search(r'(Symptoms?)[\s:\-]+(.+?)(?=(Parts to Replace|Brands|$))', remaining_text, re.IGNORECASE)
            if sym_match:
                if not fc_match and not why_match:  # Only update title if we haven't already
                    title_text = remaining_text[:sym_match.start()].strip()
                # Symptoms on same line - this is the header, actual symptoms follow on next lines
                # Don't extract symptoms content here, wait for next lines
                remaining_text = remaining_text[sym_match.end():]

            issue = {
                'title': title_text,
                'fault_codes': fault_codes_inline,
                'why': why_inline,
                'symptoms': [],
                'parts': [],
                'brands': []
            }
            
            # Helper to extract part details
            def extract_part_from_text(text, para_obj=None):
                """
                Extract part name, link, and description from text.
                Uses hyperlinks from the paragraph object if available.
                """
                part_name = text
                link = ''
                description = ''
                
                # First, try to extract hyperlinks from the paragraph object
                if para_obj:
                    hyperlinks = extract_hyperlinks_from_paragraph(para_obj, doc)
                    
                    # Filter hyperlinks to exclude those that are full URLs (the ones in parentheses)
                    # We want the hyperlink where the text is the part name (underlined text)
                    valid_hyperlinks = [(hl_text, url) for hl_text, url in hyperlinks 
                                       if not hl_text.startswith('http://') and not hl_text.startswith('https://')]
                    
                    if valid_hyperlinks:
                        # Use the first valid hyperlink (the underlined part name)
                        hyperlink_text, hyperlink_url = valid_hyperlinks[0]
                        
                        # The hyperlink text is our part name
                        part_name = hyperlink_text
                        link = hyperlink_url
                        
                        # Everything after the hyperlink text (and removing URL in parentheses) is description
                        # Remove the URL in parentheses if present
                        text_clean = text
                        url_in_parens = re.search(r'\s*\(\s*https?://[^\)]+\s*\)', text_clean)
                        if url_in_parens:
                            text_clean = text_clean[:url_in_parens.start()] + text_clean[url_in_parens.end():]
                        
                        # Find where the part name appears in the text and get everything after it
                        if hyperlink_text in text_clean:
                            idx = text_clean.index(hyperlink_text)
                            description = text_clean[idx + len(hyperlink_text):].strip()
                        
                        return {
                            'name': part_name,
                            'description': ' ' + description if description else '',
                            'link': link
                        }
                
                # Fallback: Old format handling if no hyperlinks found
                # New format: "Part Name ( url ) description"
                # Extract URL if it's in parentheses
                url_match = re.search(r'\s*\(\s*(https?://[^\)]+)\s*\)', text)
                if url_match:
                    link = url_match.group(1).strip()
                    # Everything before '(' is the part name
                    part_name = text[:url_match.start()].strip()
                    # Everything after ')' is the description
                    description = text[url_match.end():].strip()
                    if description:
                        description = ' ' + description
                else:
                    # Old format handling
                    if ' is a ' in text:
                        part_name = text.split(' is a ')[0].strip()
                        description = ' is a ' + text.split(' is a ', 1)[1].strip()
                    elif ' is an ' in text:
                        part_name = text.split(' is an ')[0].strip()
                        description = ' is an ' + text.split(' is an ', 1)[1].strip()
                    
                    # Fix for description starting with "The" inside part_name
                    if ' The ' in part_name:
                        parts = part_name.split(' The ', 1)
                        part_name = parts[0].strip()
                        description = ' The ' + parts[1].strip() + description

                    # User rule: "if last character is number or capital letter till then you have to consider model link"
                    # Scan backwards for the first char that is a digit or uppercase letter.
                    cut_idx = -1
                    for i in range(len(part_name) - 1, -1, -1):
                        if part_name[i].isdigit() or part_name[i].isupper():
                            cut_idx = i
                            break
                    
                    if cut_idx != -1:
                        # Everything after cut_idx is description/suffix
                        suffix = part_name[cut_idx+1:]
                        part_name = part_name[:cut_idx+1]
                        description = suffix + description

                    # Simple search query generation for fallback link
                    search_query = re.sub(r'[^a-zA-Z0-9\s]', '', part_name).strip()
                    link = f'https://eeuroparts.com/parts/search?q={search_query}'
                
                return {
                    'name': part_name,
                    'description': description,
                    'link': link
                }

            # Advance to next lines to find fields - SEQUENTIAL PARSING
            # Fields must appear in this order: Fault Codes -> Why it happens -> Symptoms -> Parts to Replace -> Brands
            # Track which fields have been parsed to prevent re-parsing
            fields_parsed = {
                'fault_codes': False,
                'why': False,
                'symptoms': False,
                'parts': False,
                'brands': False
            }
            
            j = i + 1
            while j < len(paragraphs):
                sub_p = paragraphs[j].strip()
                
                # Check if we hit a new category
                is_next_cat = False
                sub_p_clean = sub_p.strip().lower().rstrip(':')
                if sub_p_clean in category_keys_lower:
                    is_next_cat = True
                
                if is_next_cat:
                    break
                
                # Check for fields using regex - but only if not already parsed
                is_keyword = False
                
                # Fault Codes - only parse if not already done
                if not fields_parsed['fault_codes']:
                    fc_match = re.match(r'^(Fault Codes?|Fault Code)[\s:\-]+(.+?)(?=(Why it happens|Symptoms|Parts to Replace|Brands|$))', sub_p, re.IGNORECASE | re.DOTALL)
                    if fc_match:
                        is_keyword = True
                        fields_parsed['fault_codes'] = True
                        val = fc_match.group(2).strip()
                            
                        if val.lower() not in ['n/a', 'none', 'null', '']:
                            issue['fault_codes'] = extract_fault_codes(val)
                        
                        # Check if Why it happens is on the same line
                        remaining = sub_p[fc_match.end():]
                        why_match = re.search(r'(Why it happens)[\s:\-]+(.+?)(?=(Symptoms|Parts to Replace|Brands|$))', remaining, re.IGNORECASE | re.DOTALL)
                        if why_match:
                            fields_parsed['why'] = True
                            issue['why'] = why_match.group(2).strip()
                
                # Why it happens - only parse if fault codes already parsed (or skipped) and why not parsed yet
                if not is_keyword and not fields_parsed['why']:
                    why_match = re.match(r'^(Why it happens)[\s:\-]+(.+?)(?=(Symptoms|Parts to Replace|Brands|$))', sub_p, re.IGNORECASE | re.DOTALL)
                    if why_match:
                        is_keyword = True
                        fields_parsed['why'] = True
                        # Mark fault_codes as parsed even if we didn't find it (we're past it now)
                        fields_parsed['fault_codes'] = True
                        val = why_match.group(2).strip()
                        issue['why'] = val

                # Symptoms - only parse if we're past fault codes and why (they're done or skipped)
                if not is_keyword and not fields_parsed['symptoms']:
                    sym_match = re.match(r'^(Symptoms?)[\s:\-]*', sub_p, re.IGNORECASE)
                    if sym_match:
                        is_keyword = True
                        fields_parsed['symptoms'] = True
                        fields_parsed['fault_codes'] = True  # Mark earlier fields as done
                        fields_parsed['why'] = True
                        
                        sym_content_match = re.match(r'^(Symptoms?)[\s:\-]+(.+?)$', sub_p, re.IGNORECASE | re.DOTALL)
                        
                        # If there's content after "Symptoms:", it might be a symptom or might be empty
                        if sym_content_match:
                            sym_text = sym_content_match.group(2).strip()
                            if sym_text and not re.match(r'^(Parts to Replace|Brands)', sym_text, re.IGNORECASE):
                                # Remove numbering and check for duplicates
                                clean_sym = re.sub(r'^\d+[\.\)]\s*', '', sym_text).strip()
                                # Check if this symptom already exists (case-insensitive)
                                exists = any(re.sub(r'^\d+[\.\)]\s*', '', s).strip().lower() == clean_sym.lower() for s in issue['symptoms'])
                                if not exists:
                                    issue['symptoms'].append(sym_text)
                        
                        # Capture subsequent lines as symptoms - ONLY stop for Parts to Replace or Brands
                        k = j + 1
                        while k < len(paragraphs):
                            next_sub = paragraphs[k].strip()
                            # Check for category header
                            next_sub_clean = next_sub.strip().lower().rstrip(':')
                            is_cat_header = next_sub_clean in category_keys_lower

                            # ONLY stop for Parts to Replace or Brands (next fields in sequence)
                            # Do NOT stop for "Fault Codes" or "Why it happens" - they're part of symptoms content
                            if (re.match(r'^(Parts to Replace|Brands)[\s:\-]*', next_sub, re.IGNORECASE) or 
                                is_cat_header):
                                break
                            if next_sub:
                                # Remove numbering and check for duplicates
                                clean_next = re.sub(r'^\d+[\.\)]\s*', '', next_sub).strip()
                                # Check if this symptom already exists (case-insensitive)
                                exists = any(re.sub(r'^\d+[\.\)]\s*', '', s).strip().lower() == clean_next.lower() for s in issue['symptoms'])
                                if not exists:
                                    issue['symptoms'].append(next_sub)
                            k += 1
                        j = k - 1

                # Parts to Replace - only parse if we're past symptoms
                if not is_keyword and not fields_parsed['parts']:
                    parts_match = re.match(r'^(Parts to Replace)[\s:\-]*', sub_p, re.IGNORECASE)
                    if parts_match:
                        is_keyword = True
                        fields_parsed['parts'] = True
                        fields_parsed['fault_codes'] = True  # Mark all earlier fields as done
                        fields_parsed['why'] = True
                        fields_parsed['symptoms'] = True
                        
                        parts_content_match = re.match(r'^(Parts to Replace)[\s:\-]+(.+?)(?=(Brands|$))', sub_p, re.IGNORECASE | re.DOTALL)
                        
                        # Check if there's content on the same line
                        if parts_content_match:
                            part_text = parts_content_match.group(2).strip()
                            if part_text and not re.match(r'^(Brands)', part_text, re.IGNORECASE):
                                issue['parts'].append(extract_part_from_text(part_text, paragraph_objects[j] if j < len(paragraph_objects) else None))
                        
                            # Check if Brands is on the same line
                            remaining = sub_p[parts_content_match.end():]
                            brands_match = re.search(r'(Brands)[\s:\-]+(.+?)$', remaining, re.IGNORECASE | re.DOTALL)
                            if brands_match:
                                fields_parsed['brands'] = True
                                brands_text = brands_match.group(2).strip()
                                if 'eEuroparts Advantage:' in brands_text:
                                    brands_text = brands_text.split('eEuroparts Advantage:')[0].strip()
                                if brands_text:
                                    brand_list = [b.strip() for b in brands_text.replace(' and ', ',').split(',') if b.strip()]
                                    issue['brands'] = [{'name': b, 'link': f"https://eeuroparts.com/{b.replace(' ', '-')}"} for b in brand_list]
                        
                        # Capture subsequent part lines - ONLY stop for Brands
                        k = j + 1
                        while k < len(paragraphs):
                            next_sub = paragraphs[k].strip()
                            # Check for category header
                            next_sub_clean = next_sub.strip().lower().rstrip(':')
                            is_cat_header = next_sub_clean in category_keys_lower

                            # ONLY stop for Brands (next field in sequence)
                            if (re.match(r'^(Brands)[\s:\-]*', next_sub, re.IGNORECASE) or 
                                is_cat_header):
                                break
                            if next_sub:
                                issue['parts'].append(extract_part_from_text(next_sub, paragraph_objects[k] if k < len(paragraph_objects) else None))
                            k += 1
                        j = k - 1

                # Brands - only parse if we're past parts
                if not is_keyword and not fields_parsed['brands']:
                    brands_match = re.match(r'^(Brands)[\s:\-]*\s*(.*)', sub_p, re.IGNORECASE)
                    if brands_match:
                        is_keyword = True
                        fields_parsed['brands'] = True
                        fields_parsed['fault_codes'] = True  # Mark all earlier fields as done
                        fields_parsed['why'] = True
                        fields_parsed['symptoms'] = True
                        fields_parsed['parts'] = True
                        
                        brands_text = brands_match.group(2).strip()
                        if 'eEuroparts Advantage:' in brands_text:
                            brands_text = brands_text.split('eEuroparts Advantage:')[0].strip()
                        if brands_text:
                            brand_list = [b.strip() for b in brands_text.replace(' and ', ',').split(',') if b.strip()]
                            issue['brands'] = [{'name': b, 'link': f"https://eeuroparts.com/{b.replace(' ', '-')}"} for b in brand_list]

                # Implicit Symptoms (lines with colons that aren't keywords)
                # Only treat as symptoms if we haven't reached Parts to Replace yet
                if not is_keyword and ':' in sub_p and issue['title'] and not fields_parsed['parts']:
                    parts = sub_p.split(':', 1)
                    key_candidate = parts[0].strip()
                    # Heuristic: if key is short and not a sentence, treat as symptom
                    # Also ensure it doesn't look like a numbered list item (which could be a new issue title)
                    if (len(key_candidate) < 50 and 
                        not any(x in key_candidate.lower() for x in ['note', 'important']) and
                        not re.match(r'^\d+[\.\)]', sub_p)):
                        
                        is_keyword = True
                        fields_parsed['symptoms'] = True
                        fields_parsed['fault_codes'] = True
                        fields_parsed['why'] = True
                        
                        # Remove numbering and check for duplicates before adding
                        clean_sub = re.sub(r'^\d+[\.\)]\s*', '', sub_p).strip()
                        exists = any(re.sub(r'^\d+[\.\)]\s*', '', s).strip().lower() == clean_sub.lower() for s in issue['symptoms'])
                        if not exists:
                            issue['symptoms'].append(sub_p)
                        
                        # Capture subsequent lines - ONLY stop for Parts to Replace or Brands
                        k = j + 1
                        while k < len(paragraphs):
                            next_sub = paragraphs[k].strip()
                            next_sub_clean = next_sub.strip().lower().rstrip(':')
                            is_cat_header = next_sub_clean in category_keys_lower

                            # ONLY stop for Parts to Replace or Brands (next fields in sequence)
                            if (re.match(r'^(Parts to Replace|Brands)[\s:\-]*', next_sub, re.IGNORECASE) or 
                                is_cat_header or
                                re.match(r'^\d+[\.\)]', next_sub)): # Stop if next line looks like a numbered issue
                                break
                            if next_sub:
                                # Remove numbering and check for duplicates
                                clean_next = re.sub(r'^\d+[\.\)]\s*', '', next_sub).strip()
                                exists = any(re.sub(r'^\d+[\.\)]\s*', '', s).strip().lower() == clean_next.lower() for s in issue['symptoms'])
                                if not exists:
                                    issue['symptoms'].append(next_sub)
                            k += 1
                        j = k - 1
                
                if not is_keyword:
                    # Line is not a keyword.
                    # If ALL fields are parsed (brands is the last one), we've completed the issue
                    if fields_parsed['brands']:
                        break
                    # If we have already collected some fields and encounter a non-keyword line
                    # that doesn't match our parsing rules, it's likely a NEW ISSUE TITLE
                    elif issue['why'] or issue['symptoms'] or issue['parts']:
                        break
                    else:
                        # If we haven't collected fields yet, it might be part of the title or description?
                        # Or maybe the issue title spans multiple lines?
                        # Let's append to title if it's not too long
                        if len(issue['title']) < 200:
                            issue['title'] += " " + sub_p
                
                j += 1
            
            # Add issue to category
            # Clean title
            issue['title'] = re.sub(r'^\d+[\.\)]\s*', '', issue['title']).strip()
            if issue['title']:
                data['issues'][current_category].append(issue)
            
            # Move main index
            i = j
            continue

        i += 1
    
    return data

def generate_html(data, template_path, output_path):
    with open(template_path, 'r', encoding='utf-8') as f:
        template_str = f.read()
    template = Template(template_str)
    html = template.render(**data)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

if __name__ == '__main__':
    import os
    import sys
    import glob
    
    template_path = 'template.html'
    
    # Check if a specific file is provided as an argument
    if len(sys.argv) > 1:
        docx_path = sys.argv[1]
        if not os.path.exists(docx_path):
            print(f"Error: File '{docx_path}' not found.")
            sys.exit(1)
            
        # User requested output.html for single file processing
        output_path = 'output.html'
        
        print(f"Processing {docx_path}...")
        try:
            data = parse_word_document(docx_path)
            generate_html(data, template_path, output_path)
            print(f'HTML generated successfully: {output_path}')
            print(f'Vehicle: {data["vehicle_heading"]}')
            if data["description_text"]:
                print(f'Description: {data["description_text"][:100]}...')
            else:
                print('Description: Not found')
            print(f'Specs: {len(data["specs"])} categories')
            print(f'Issues: {sum(len(v) for v in data["issues"].values())} total')
        except Exception as e:
            print(f"Error processing {docx_path}: {e}")
            
    else:
        # Batch mode - process all files in directory
        print("No input file provided. Scanning directory for .docx files...")
        docx_files = glob.glob('*.docx')
        
        if not docx_files:
            print("No .docx files found.")
        
        for docx_path in docx_files:
            print(f"Processing {docx_path}...")
            output_path = os.path.splitext(docx_path)[0] + '.html'
            
            try:
                data = parse_word_document(docx_path)
                generate_html(data, template_path, output_path)
                print(f'HTML generated successfully: {output_path}')
                print(f'Vehicle: {data["vehicle_heading"]}')
                if data["description_text"]:
                    print(f'Description: {data["description_text"][:100]}...')
                else:
                    print('Description: Not found')
                print(f'Specs: {len(data["specs"])} categories')
                print(f'Issues: {sum(len(v) for v in data["issues"].values())} total')
                print('-' * 40)
            except Exception as e:
                print(f"Error processing {docx_path}: {e}")

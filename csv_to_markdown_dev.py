import csv
import os
import requests
from io import StringIO
import re
import shutil
import itertools
from concurrent.futures import ThreadPoolExecutor
import time
from functools import lru_cache
import logging
from datetime import datetime
from collections import defaultdict

# Global cache for note contents
note_content_cache = {}

# Initialize link_dict as a global variable
link_dict = {}

def cleanup_vault(vault_path):
    print(f"Cleaning up vault at: {vault_path}")
    for item in os.listdir(vault_path):
        item_path = os.path.join(vault_path, item)
        
        if item == ".obsidian":
            print(f"Skipping .obsidian folder: {item_path}")
            continue
        
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path):
                os.unlink(item_path)
                print(f"Deleted file: {item_path}")
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
                print(f"Deleted folder: {item_path}")
        except Exception as e:
            print(f"Error deleting {item_path}: {e}")

vault_path = r"G:\My Drive\Drive\Gaming_Music_Comics_software\Weather Factory\Obsidian"

# Set up logging at the start of your script (right after imports)
log_file = os.path.join(vault_path, "obsidian_import_log.txt")
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filemode='w'  # Overwrite existing log
)
logger = logging.getLogger()

book_of_hours_sheets = {
    'Memories': '57430724',
    'Consider Masterlist': '432406626',
    'BoH Books': '161140297',
    'Skills': '887548006',
    'Wisdom Branch': '1238682685',
    'Sustenance': '652283921',
    'Inedible Cooking Ingredient': '126757897',
    'Beverages': '367912105',
    'Candles': '745028529',
    'Flowers': '2030685183',
    'Pets': '1839755106',
    'Explorable': '203136705',
    'Purchasable': '484734701',
    'Crafted Product': '1675846500',
    'Crafting Recipe': '1760695989',
    'Rooms': '1889312489',
    'Workstations': '1351975106',
    'Visitors': '1721650548',
    'Assistance': '1059413647',
    'Exploration': '1937107873',
    'Incidents': '1163393403',
    'History': '1252134190',
    'Conversations': '605494400',
    'Hours': '1981936845',
    'Weather': '1191136372',
    'Principle': '1007960833',
    'Soul Card': '338064931',
    'Keywords': '1084909450',
    'Player Librarians': '1610034163',
    'History Victory': '1427507298',
    'Lighthouse Institute Victory': '873533606',
}

boh_link = "https://docs.google.com/spreadsheets/d/1p1aWr5N0GXdATgP9jM9pB_lL5HHsKEaQDU8zp0h7GMM/edit?gid=dict_url_reference#gid=dict_url_reference"

subfolders_dict = {
    'Book of Hours': {
        'folder_name': 'Book of Hours',
        'sheets': book_of_hours_sheets,
        'link_template': boh_link
    }
}

keyword_sheets = ["Keywords", "Glossary"]

def generate_sheet_urls(base_url, sheets_dict):
    sheet_urls = []
    for sheet_name, gid in sheets_dict.items():
        sheet_url = base_url.replace("dict_url_reference", gid)
        sheet_urls.append(sheet_url)
        print(f"Generated sheet URL for {sheet_name}: {sheet_url}")
    return sheet_urls

def extract_gid(url):
    print(f"Extracting gid from URL: {url}")
    match = re.search(r"gid=(\d+)", url)
    if match:
        gid = match.group(1)
        print(f"Extracted gid: {gid}")
        return gid
    print("Warning: No gid found in the URL.")
    return None

processed_data = {}

for subfolder_key, subfolder_data in subfolders_dict.items():
    subfolder_path = os.path.join(vault_path, subfolder_data['folder_name'])
    os.makedirs(subfolder_path, exist_ok=True)
    
    sheet_urls = generate_sheet_urls(
        subfolder_data['link_template'],
        subfolder_data['sheets']
    )
    
    processed_data[subfolder_key] = {
        'folder_name': subfolder_data['folder_name'],
        'vault_path': subfolder_path,
        'sheet_urls': sheet_urls,
        'csv_urls': []
    }
    
    for url in sheet_urls:
        gid = extract_gid(url)
        if gid:
            csv_url = f"https://docs.google.com/spreadsheets/d/1p1aWr5N0GXdATgP9jM9pB_lL5HHsKEaQDU8zp0h7GMM/export?format=csv&gid={gid}"
            processed_data[subfolder_key]['csv_urls'].append(csv_url)
            print(f"Generated CSV export URL: {csv_url}")

delimiters = [",", ":", "_", "-", " "]
excluded_words = {"the", "a", "an", "and", "or", "of", "in", "to", "for", "with", "on", "at", "by", "as"}
excluded_digits = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}

priority_link_references = set()
secondary_link_references = set()

def split_value(value):
    parts = [value]
    for delimiter in delimiters:
        new_parts = []
        for part in parts:
            new_parts.extend(part.split(delimiter))
        parts = new_parts
    return [part.strip() for part in parts if part.strip()]

def download_csv(url):
    print(f"Downloading CSV from URL: {url}")
    response = requests.get(url)
    response.raise_for_status()
    print("CSV downloaded successfully.")
    return StringIO(response.text)

def escape_markdown(text):
    chars_to_escape = {'\\', '#', '^', '|', '{', '}'}
    escaped_text = []
    for char in text:
        if char in chars_to_escape:
            escaped_text.append('\\' + char)
        else:
            escaped_text.append(char)
    return ''.join(escaped_text)

def sanitize_value(value):
    value = value.strip()
    value = escape_markdown(value)
    return value

def sanitize_filename(filename):
    filename = filename.replace(':', '_')
    filename = re.sub(r'[\\/*?:"<>|]', '_', filename)
    return filename

def sanitize_link(link):
    return link.replace(':', '_')

def create_link_references(csv_data, sheet_name, subfolder_name):
    print(f"Creating link references for sheet: {sheet_name} in {subfolder_name}...")
    try:
        reader = csv.DictReader(csv_data)
        if not reader.fieldnames:
            print("Warning: No headers found in the CSV file.")
            return
        
        folder_name = sanitize_value(reader.fieldnames[0])
        print(f"Using folder name: {folder_name}")
        
        is_keywords_sheet = (sheet_name == "Keywords")
        
        csv_data.seek(0)
        next(reader)
        for row in reader:
            if reader.fieldnames and row:
                first_column_value = row[reader.fieldnames[0]]
                if first_column_value.strip():
                    sanitized_value = sanitize_value(first_column_value)
                    full_reference = f"{subfolder_name}/{folder_name}/{sanitized_value}"
                    priority_link_references.add(full_reference)
                    print(f"Added to priority link references: {full_reference}")
                    
                    if is_keywords_sheet:
                        for header in reader.fieldnames[1:]:
                            if header in row and row[header].strip():
                                sanitized_cell_value = sanitize_value(row[header])
                                full_cell_reference = f"{subfolder_name}/{folder_name}/{sanitized_cell_value}"
                                priority_link_references.add(full_cell_reference)
                                print(f"Added Keywords sheet value to priority links: {full_cell_reference}")
    except Exception as e:
        print(f"Error processing CSV: {e}")

def sanitize_cell_value(value):
    if not value:
        return ""
    value = ' '.join(value.splitlines())
    if ':' in value:
        value = value.replace(': ', '：')
    return value.strip()

def get_apostrophe_variants(text):
    variants = {text}
    
    if "'" in text:
        variants.add(text.replace("'", ""))
        variants.add(text.replace("'", " "))
    
    if text.endswith("'s"):
        base = text[:-2]
        variants.update({
            base,
            base + "s",
            base + "'",
            base + "s'"
        })
    elif text.endswith("s'"):
        base = text[:-1]
        variants.update({
            base,
            base + "s",
            base + "'s",
            base[:-1]
        })
    elif text.endswith("s"):
        variants.update({
            text + "'",
            text + "'s",
            text[:-1] + "'",
            text[:-1] + "'s"
        })
    
    return variants

def process_csv(csv_data, sheet_index, subfolder_key, keyword_sheets):
    """Process CSV data with proper sheet identification"""
    try:
        # Get the actual sheet name from our configuration
        sheet_name = list(subfolders_dict[subfolder_key]['sheets'].keys())[sheet_index]
        logger.info(f"Starting to process sheet: {sheet_name} (index: {sheet_index})")
        
        csv_data.seek(0)
        reader = csv.reader(csv_data)
        headers = [h.strip() for h in next(reader)]
        
        if not headers:
            logger.warning(f"Empty sheet {sheet_name}")
            return
            
        logger.info(f"Headers found: {headers}")
        
        if sheet_name in keyword_sheets:
            logger.info("Processing as keyword sheet")
            process_keywords_sheet(csv_data, reader, headers, subfolder_key, processed_data[subfolder_key])
        else:
            logger.info("Processing as normal sheet")
            process_normal_sheet(csv_data, sheet_index, subfolder_key, processed_data[subfolder_key])
            
    except Exception as e:
        logger.error(f"Error processing sheet {sheet_index}: {str(e)}", exc_info=True)
        raise

def process_keywords_sheet(csv_data, reader, headers, subfolder_key, subfolder_data):
    print("Processing Keywords sheet with column-based subfolders")
    
    keywords_base_folder = os.path.join(subfolder_data['vault_path'], "Keywords")
    os.makedirs(keywords_base_folder, exist_ok=True)
    
    if 'sheet_folders' not in processed_data[subfolder_key]:
        processed_data[subfolder_key]['sheet_folders'] = {}
    
    for col_index, header in enumerate(headers):
        if not header:
            continue
            
        process_keywords_column(csv_data, col_index, header, subfolder_key, keywords_base_folder)

def process_keywords_column(csv_data, col_index, header, subfolder_key, base_folder):
    header_folder_name = sanitize_value(header).replace(':', '_')
    header_folder = os.path.join(base_folder, header_folder_name)
    os.makedirs(header_folder, exist_ok=True)
    
    csv_data.seek(0)
    reader = csv.reader(csv_data)
    next(reader)
    
    processed_values = set()
    
    for row in reader:
        if len(row) > col_index and row[col_index].strip():
            value = row[col_index].strip()
            if value not in processed_values:
                processed_values.add(value)
                create_keyword_file(value, header_folder_name, subfolder_key, header_folder)
                
                filename_value = sanitize_value(value).replace(':', '_')
                full_reference = f"{subfolder_key}/Keywords/{header_folder_name}/{filename_value}"
                priority_link_references.add(full_reference)
                print(f"Added to priority links: {full_reference}")
    
    folder_key = f"Keywords/{header_folder_name}"
    processed_data[subfolder_key]['sheet_folders'][folder_key] = {
        'items': [v.replace(':', '_') for v in processed_values],
        'path': header_folder
    }

def create_keyword_file(value, header_folder_name, subfolder_key, header_folder):
    filename_value = sanitize_value(value).replace(':', '_')
    base_filename = sanitize_filename(filename_value)
    filename = f"{base_filename}.md"
    filepath = os.path.join(header_folder, filename)
    
    linked_notes = find_notes_referencing_keyword(value)
    
    with open(filepath, 'w', encoding='utf-8') as md_file:
        if linked_notes:
            md_file.write("## Linked Notes\n")
            for note in sorted(linked_notes):
                md_file.write(f"- [[{note}]]\n")
    
    print(f"Created keyword file with {len(linked_notes)} links: {filepath}")

def find_notes_referencing_keyword(keyword_value):
    linked_notes = set()
    sanitized_keyword = sanitize_value(keyword_value).replace(':', '_')
    
    cached_notes = get_cached_notes()
    
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = []
        for note_path, note_content in cached_notes:
            futures.append(
                executor.submit(
                    check_note_for_keyword,
                    note_path,
                    note_content,
                    sanitized_keyword
                )
            )
        
        for future in futures:
            result = future.result()
            if result:
                linked_notes.add(result)
    
    return linked_notes

def check_note_for_keyword(note_path, note_content, sanitized_keyword):
    if "Keywords/" in note_path:
        return None
        
    for variant in get_apostrophe_variants(sanitized_keyword):
        if re.search(r'\b' + re.escape(variant) + r'\b', note_content, re.IGNORECASE):
            return convert_path_to_reference(note_path)
    return None

def convert_path_to_reference(filepath):
    rel_path = os.relpath(filepath, start=vault_path)
    ref_path = rel_path[:-3].replace('\\', '/')
    subfolder_key = next(k for k in processed_data if filepath.startswith(processed_data[k]['vault_path']))
    return f"{subfolder_key}/{ref_path}"

def get_cached_notes():
    if not note_content_cache:
        print("Building note content cache...")
        start_time = time.time()
        
        with ThreadPoolExecutor(max_workers=8) as executor:
            futures = []
            for subfolder_key, subfolder_data in processed_data.items():
                vault_path = subfolder_data['vault_path']
                for root, _, files in os.walk(vault_path):
                    for file in files:
                        if file.endswith('.md'):
                            filepath = os.path.join(root, file)
                            futures.append(executor.submit(read_note_content, filepath))
            
            for future in futures:
                filepath, content = future.result()
                note_content_cache[filepath] = content
        
        print(f"Cache built in {time.time() - start_time:.2f} seconds")
    
    return note_content_cache.items()

def read_note_content(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return (filepath, f.read())
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return (filepath, "")

def process_normal_sheet(csv_data, sheet_index, subfolder_key, subfolder_data):
    """Process a sheet with proper sheet name identification"""
    # Get the actual sheet name from our configuration
    sheet_name = list(subfolders_dict[subfolder_key]['sheets'].keys())[sheet_index]
    logger.info(f"Normal processing for sheet: {sheet_name}")
    
    # SPECIAL CASE: History sheet - check directly by name
    if sheet_name == "History":
        logger.info("Identified as History sheet - using special processor")
        process_history_sheet(csv_data, subfolder_key, subfolder_data)
        return
    
    # Normal sheet processing
    csv_data.seek(0)
    reader = csv.reader(csv_data)
    headers = [h.strip() for h in next(reader)]
    
    folder_name = sanitize_value(headers[0]).replace(':', '_')
    logger.info(f"Using folder name: {folder_name} for sheet: {sheet_name}")
    
    sheet_folder = os.path.join(subfolder_data['vault_path'], folder_name)
    os.makedirs(sheet_folder, exist_ok=True)
    
    if 'sheet_folders' not in processed_data[subfolder_key]:
        processed_data[subfolder_key]['sheet_folders'] = {}
    processed_data[subfolder_key]['sheet_folders'][folder_name] = {
        'items': [],
        'path': sheet_folder
    }
    
    # Process headers
    header_count = {}
    sanitized_headers = []
    for header in headers:
        sanitized_header = sanitize_value(header)
        if sanitized_header in header_count:
            header_count[sanitized_header] += 1
            sanitized_headers.append(f"{sanitized_header}_{header_count[sanitized_header]}")
        else:
            header_count[sanitized_header] = 1
            sanitized_headers.append(sanitized_header)
    
    logger.debug(f"Sanitized headers: {sanitized_headers}")
    
    # Process rows
    csv_data.seek(0)
    next(reader)  # Skip header
    for row_idx, row in enumerate(reader, 1):
        if not row:
            logger.debug(f"Skipping empty row {row_idx}")
            continue
        process_normal_row(row, sanitized_headers, folder_name, subfolder_key, sheet_folder)

def process_history_year_entries(year_value, entries, headers, folder_name, subfolder_key, sheet_folder):
    """Process all entries for a year with verification"""
    filename_value = sanitize_value(year_value).replace(':', '_')
    base_filename = sanitize_filename(filename_value)
    filename = f"{base_filename}.md"
    filepath = os.path.join(sheet_folder, filename)
    
    # Verify we have all entries
    print(f"Creating file for {year_value} with {len(entries)} entries")
    for i, entry in enumerate(entries, 1):
        print(f"  Entry {i}: {entry['row'][0:3]}...")  # Print first few columns
        
    # Prepare content
    front_matter = f"---\nYear: {year_value}\n---\n\n"
    entries_content = "## Historical Entries\n\n"
    links = set()
    
    for entry in entries:
        row = entry['row']
        entry_num = entry['index']
        
        entries_content += f"### Entry {entry_num}\n"
        for col_idx, value in enumerate(row):
            if value and value.strip():
                header = headers[col_idx]
                entries_content += f"- **{header}**: {sanitize_cell_value(value)}\n"
                
                # Process links
                if header != "Year":  # Skip year field
                    sanitized_value = sanitize_value(value).replace(':', '_')
                    for ref in priority_link_references:
                        if ref.endswith(sanitized_value) and ref != f"{subfolder_key}/{folder_name}/{filename_value}":
                            links.add(f"[[{ref}]]")
        entries_content += "\n"
    
    # Write to file
    with open(filepath, 'w', encoding='utf-8') as md_file:
        md_file.write(front_matter)
        md_file.write(entries_content)
        
        if links:
            md_file.write("## Links\n")
            for link in sorted(links):
                md_file.write(f"{link}\n")
    
    print(f"Successfully created file with {len(entries)} entries: {filepath}")

def process_history_sheet(csv_data, subfolder_key, subfolder_data):
    """Special processing for History sheet with detailed logging"""
    logger.info("=== PROCESSING HISTORY SHEET ===")
    
    # Read ALL data at once
    csv_data.seek(0)
    all_rows = list(csv.reader(csv_data))
    headers = [h.strip() for h in all_rows[0]]
    data_rows = all_rows[1:]
    
    logger.info(f"Found {len(data_rows)} total rows (excluding header)")
    logger.info(f"Headers: {headers}")

    folder_name = "History"
    sheet_folder = os.path.join(subfolder_data['vault_path'], folder_name)
    os.makedirs(sheet_folder, exist_ok=True)
    
    # Initialize tracking
    processed_data[subfolder_key]['sheet_folders'][folder_name] = {
        'items': [],
        'path': sheet_folder
    }
    
    # Process headers
    sanitized_headers = []
    header_count = {}
    for header in headers:
        sanitized_header = sanitize_value(header)
        if sanitized_header in header_count:
            header_count[sanitized_header] += 1
            sanitized_headers.append(f"{sanitized_header}_{header_count[sanitized_header]}")
        else:
            header_count[sanitized_header] = 1
            sanitized_headers.append(sanitized_header)
    
    logger.info(f"Sanitized headers: {sanitized_headers}")
    
    # Group rows by year
    year_entries = defaultdict(list)
    for row_idx, row in enumerate(data_rows, 1):
        if not row or not row[0].strip():
            logger.warning(f"Skipping empty row {row_idx}")
            continue
            
        year_value = sanitize_value(row[0].strip())
        year_entries[year_value].append(row)
        if year_value not in processed_data[subfolder_key]['sheet_folders'][folder_name]['items']:
            processed_data[subfolder_key]['sheet_folders'][folder_name]['items'].append(year_value)
        logger.debug(f"Row {row_idx} assigned to year: {year_value}")

    # Log year distribution
    logger.info(f"Year distribution: { {k: len(v) for k, v in year_entries.items()} }")
    
    # Process each year
    for year_value, entries in year_entries.items():
        logger.info(f"Processing {len(entries)} entries for year: {year_value}")
        
        filename = f"{sanitize_filename(year_value.replace(':', '_'))}.md"
        filepath = os.path.join(sheet_folder, filename)
        
        # Build content
        content = f"---\nYear: {year_value}\n---\n\n## Historical Entries\n\n"
        links = set()
        
        for entry_num, row in enumerate(entries, 1):
            logger.debug(f"  Processing entry {entry_num} for {year_value}")
            content += f"### Entry {entry_num}\n"
            
            for col_idx, value in enumerate(row):
                if value and value.strip():
                    header = sanitized_headers[col_idx]
                    safe_value = sanitize_cell_value(value)
                    content += f"- **{header}**: {safe_value}\n"
                    
                    # Log link processing
                    if header != "Year":
                        for ref in priority_link_references:
                            if ref.endswith(safe_value.replace(':', '_')):
                                logger.debug(f"    Found link: {ref}")
                                links.add(f"[[{ref}]]")
            
            content += "\n"
        
        # Add links section
        if links:
            logger.info(f"  Found {len(links)} links for {year_value}")
            content += "## Links\n"
            for link in sorted(links):
                content += f"{link}\n"
        
        # Write file
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        
        logger.info(f"Created file: {filename} with {len(entries)} entries")

def process_history_year(year_value, entries, headers, folder_name, subfolder_key, sheet_folder):
    """Process all entries for a single year and combine them into one file"""
    filename_value = sanitize_value(year_value).replace(':', '_')
    base_filename = sanitize_filename(filename_value)
    filename = f"{base_filename}.md"
    filepath = os.path.join(sheet_folder, filename)
    
    # Prepare front matter with the year
    front_matter = f"---\nYear: {year_value}\n---\n\n"
    
    # Prepare entries content
    entries_content = "## Historical Entries\n\n"
    links = set()
    
    # Process each entry
    for entry_idx, row in enumerate(entries, 1):
        entries_content += f"### Entry {entry_idx}\n"
        
        # Process each column in the row
        for col_idx, value in enumerate(row):
            if value and value.strip():
                header = headers[col_idx]
                sanitized_value = sanitize_cell_value(value)
                entries_content += f"- **{header}**: {sanitized_value}\n"
                
                # Collect links (skip year column)
                if header != "Year":
                    for ref in priority_link_references:
                        ref_name = ref.split('/')[-1]
                        if ref_name.lower() in sanitized_value.lower():
                            links.add(f"[[{ref}]]")
        
        entries_content += "\n"
    
    # Write to file
    with open(filepath, 'w', encoding='utf-8') as md_file:
        md_file.write(front_matter)
        md_file.write(entries_content)
        
        # Write links section if we have any
        if links:
            md_file.write("## Links\n")
            for link in sorted(links):
                md_file.write(f"{link}\n")
    
    print(f"Created History file for {year_value} with {len(entries)} entries: {filepath}")

def process_normal_row(row, sanitized_headers, folder_name, subfolder_key, sheet_folder):
    filename_value = sanitize_value(row[0].strip())
    if not filename_value:
        filename_value = "Untitled"
    filename_value = filename_value.replace(':', '_')
    
    filename_variants = get_apostrophe_variants(filename_value)
    base_filename = sanitize_filename(filename_value)
    filename = f"{base_filename}.md"
    filepath = os.path.join(sheet_folder, filename)
    
    if filename not in link_dict:
        link_dict[filename] = set()

    if filename_value:
        processed_data[subfolder_key]['sheet_folders'][folder_name]['items'].append(filename_value)
    
    write_normal_markdown_file(row, sanitized_headers, filename_value, filename_variants, filepath, folder_name, subfolder_key)

def write_normal_markdown_file(row, sanitized_headers, filename_value, filename_variants, filepath, folder_name, subfolder_key):
    with open(filepath, 'w', encoding='utf-8') as md_file:
        md_file.write("---\n")
        for i, value in enumerate(row):
            if value and value.strip():
                safe_key = sanitized_headers[i]
                safe_value = sanitize_cell_value(value)
                
                if 'description' in safe_key.lower() or 'note' in safe_key.lower():
                    md_file.write(f"{safe_key}: |\n  {safe_value.replace('：', ':')}\n")
                else:
                    md_file.write(f"{safe_key}: {safe_value}\n")
        md_file.write("---\n\n## Links\n")
        
        linked_values = set()
        for i, value in enumerate(row):
            if value and value.strip():
                process_cell_for_links(value, filename_value, sanitized_headers[i], filename_variants, linked_values, folder_name, subfolder_key)
        
        for linked_value in sorted(linked_values):
            md_file.write(f"- {linked_value}\n")
    
    print(f"Created: {filepath}")

def process_cell_for_links(value, filename_value, sanitized_header, filename_variants, linked_values, folder_name, subfolder_key):
    sanitized_value = sanitize_value(value)
    if sanitized_value == filename_value:
        return  # Skip self-references
        
    for reference in priority_link_references:
        reference_parts = reference.split("/")
        reference_filename = reference_parts[-1]
        
        # Skip self-references
        if reference == f"{subfolder_key}/{folder_name}/{filename_value}":
            continue
            
        for variant in filename_variants:
            if variant == reference_filename:
                add_link(linked_values, reference)
        
        for variant in get_apostrophe_variants(sanitized_value):
            if variant == reference_filename:
                add_link(linked_values, reference)
        
        for variant in get_apostrophe_variants(sanitized_header):
            if variant == reference_filename:
                add_link(linked_values, reference)

def add_link(linked_values, reference):
    linked_values.add(f"[[{reference}]]")
    link_dict[reference] = reference

def create_masterlists(subfolder_key):
    subfolder_data = processed_data[subfolder_key]
    masterlist_folder = os.path.join(subfolder_data['vault_path'], "Masterlists")
    os.makedirs(masterlist_folder, exist_ok=True)
    
    for folder_name, folder_info in subfolder_data['sheet_folders'].items():
        if folder_name.startswith("Keywords/"):
            continue
            
        safe_filename = sanitize_filename(folder_name)
        masterlist_file = os.path.join(masterlist_folder, f"{safe_filename}.md")
        
        with open(masterlist_file, 'w', encoding='utf-8') as f:
            f.write(f"# {folder_name} Masterlist\n\n")
            for item in folder_info['items']:
                link_text = f"{subfolder_key}/{folder_name}/{item.replace(':', '_')}"
                f.write(f"- [[{link_text}]]\n")
        
        masterlist_filename = f"Masterlists/{safe_filename}.md"
        if masterlist_filename not in link_dict:
            link_dict[masterlist_filename] = set()
        
        for item in folder_info['items']:
            for variant in get_apostrophe_variants(item):
                link_dict[masterlist_filename].add(f"{subfolder_key}/{folder_name}/{variant.replace(':', '_')}")

def update_reverse_links():
    print("Updating reverse links...")
    
    for filename, links in link_dict.items():
        for link in links:
            linked_filename = link.split("/")[-1] + ".md"
            if linked_filename in link_dict:
                reverse_link = "/".join(link.split("/")[:-1] + [filename.replace(".md", "")])
                link_dict[linked_filename].add(reverse_link)
                print(f"Added reverse link: [[{reverse_link}]] to {linked_filename}")
    
    for filename, links in link_dict.items():
        subfolder_key = None
        filepath = None
        for key, data in processed_data.items():
            if filename.startswith(data['vault_path']):
                subfolder_key = key
                filepath = filename
                break
        
        if not filepath or not subfolder_key:
            print(f"Warning: Could not determine subfolder for file {filename}")
            continue
            
        if os.path.exists(filepath):
            with open(filepath, 'r+', encoding='utf-8') as md_file:
                content = md_file.read()
                md_file.seek(0)
                
                if "## Links" not in content:
                    md_file.seek(0, 2)
                    md_file.write("\n## Links\n")
                    for link in sorted(links):
                        md_file.write(f"- [[{link}]]\n")
                    print(f"Added '## Links' section to: {filepath}")
                else:
                    md_file.seek(0)
                    lines = md_file.readlines()
                    md_file.seek(0)
                    md_file.truncate()
                    
                    links_section = False
                    for line in lines:
                        if line.strip() == "## Links":
                            links_section = True
                        md_file.write(line)
                    
                    if links_section:
                        for link in sorted(links):
                            if f"[[{link}]]" not in content:
                                md_file.write(f"- [[{link}]]\n")
                        print(f"Appended links to: {filepath}")
        else:
            print(f"File not found: {filepath}")

def write_link_references():
    print("Writing link references to file...")
    for subfolder_key, subfolder_data in processed_data.items():
        link_reference_file = os.path.join(subfolder_data['vault_path'], "link_references.txt")
        with open(link_reference_file, 'w', encoding='utf-8') as f:
            f.write("Priority Link References:\n")
            for reference in sorted(priority_link_references):
                if reference.startswith(subfolder_key):
                    f.write(f"{reference}\n")
            f.write("\nSecondary Link References:\n")
            for reference in sorted(secondary_link_references):
                if reference.startswith(subfolder_key):
                    f.write(f"{reference}\n")
        print(f"Link references written to: {link_reference_file}")

def main():
    logger.info("=== SCRIPT STARTED ===")
    logger.info(f"Vault path: {vault_path}")
    
    try:
        for subfolder_key, subfolder_data in processed_data.items():
            print(f"\nProcessing subfolder: {subfolder_key}")
            
            print("Step 0: Cleaning up the vault...")
            cleanup_vault(subfolder_data['vault_path'])
            
            print("Step 1: Creating link references...")
            for i, csv_url in enumerate(subfolder_data['csv_urls']):
                try:
                    sheet_name = list(subfolders_dict[subfolder_key]['sheets'].keys())[i]
                    print(f"Processing sheet {i + 1} ({sheet_name})...")
                    csv_data = download_csv(csv_url)
                    create_link_references(csv_data, sheet_name, subfolder_key)
                    process_csv(csv_data, i, subfolder_key, keyword_sheets)
                except Exception as e:
                    print(f"Error downloading or processing CSV for sheet {i + 1}: {e}")

        print("\nStep 3: Creating masterlists...")
        for subfolder_key in processed_data:
            create_masterlists(subfolder_key)

        print("Step 4: Updating reverse links...")
        update_reverse_links()
        
        print("Step 5: Writing link references to file...")
        write_link_references()
        
        print("Script completed successfully.")
    
        logger.info("Script completed successfully")
        
    except Exception as e:
        logger.error(f"Script failed: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    main()
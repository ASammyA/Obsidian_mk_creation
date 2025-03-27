import csv
import os
import requests
from io import StringIO
import re
import shutil

# Initialize link_dict as a global variable
link_dict = {}

def cleanup_vault(vault_path):
    print(f"Cleaning up vault at: {vault_path}")
    for item in os.listdir(vault_path):
        item_path = os.path.join(vault_path, item)
        
        # Skip the .obsidian folder
        if item == ".obsidian":
            print(f"Skipping .obsidian folder: {item_path}")
            continue
        
        # Remove the item (file or folder)
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path):
                os.unlink(item_path)
                print(f"Deleted file: {item_path}")
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
                print(f"Deleted folder: {item_path}")
        except Exception as e:
            print(f"Error deleting {item_path}: {e}")

# Set your Obsidian vault directory
vault_path = r"G:\My Drive\Drive\Gaming_Music_Comics_software\Weather Factory\Obsidian" 

book_of_hours_sheets = {
    'Memories': '57430724'
    , 'Consider Masterlist': '432406626'
    , 'BoH Books': '161140297'
    , 'Skills': '887548006'
    , 'Wisdom Branch': '1238682685'
    , 'Sustenance': '652283921'
    , 'Inedible Cooking Ingredient': '126757897'
    , 'Beverages': '367912105'
    , 'Candles': '745028529'
    , 'Flowers': '2030685183'
    , 'Pets': '1839755106'
    , 'Explorable': '203136705'
    , 'Purchasable': '484734701'
    , 'Crafted Product': '1675846500'
    , 'Crafting Recipe': '1760695989'
    , 'Rooms': '1889312489'
    , 'Workstations': '1351975106'
    , 'Visitors': '1721650548'
    , 'Assistance': '1059413647'
    , 'Exploration': '1937107873'
    , 'Incidents': '1163393403'
    , 'History': '1252134190'
    , 'Conversations': '605494400'
    , 'Hours': '1981936845'
    , 'Weather': '1191136372'
    , 'Principle': '1007960833'
    , 'Soul Card': '338064931'
    , 'Keywords': '1084909450'
    , 'Player Librarians': '1610034163'
    , 'History Victory': '1427507298'
    , 'Lighthouse Institute Victory': '873533606'
}

boh_link = "https://docs.google.com/spreadsheets/d/1p1aWr5N0GXdATgP9jM9pB_lL5HHsKEaQDU8zp0h7GMM/edit?gid=dict_url_reference#gid=dict_url_reference"

def generate_sheet_urls(base_url, sheets_dict):
    """
    Generate sheet URLs from a base URL and a dictionary of sheet names and gids.
    
    Args:
        base_url (str): The base URL template containing 'dict_url_reference' to be replaced
        sheets_dict (dict): Dictionary of sheet names and their gids
        
    Returns:
        list: List of generated sheet URLs
    """
    sheet_urls = []
    for sheet_name, gid in sheets_dict.items():
        sheet_url = base_url.replace("dict_url_reference", gid)
        sheet_urls.append(sheet_url)
        print(f"Generated sheet URL for {sheet_name}: {sheet_url}")
    return sheet_urls

# delimiters to split values
delimiters = [",", ":", "_", "-", " "]

# Generic words and digits to exclude
excluded_words = {"the", "a", "an", "and", "or", "of", "in", "to", "for", "with", "on", "at", "by", "as"}
excluded_digits = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}

# Function to extract gid from a sheet URL
def extract_gid(url):
    print(f"Extracting gid from URL: {url}")
    match = re.search(r"gid=(\d+)", url)
    if match:
        gid = match.group(1)
        print(f"Extracted gid: {gid}")
        return gid
    print("Warning: No gid found in the URL.")
    return None

# Generate CSV export URLs
csv_urls = []
for url in sheet_urls:
    gid = extract_gid(url)
    if gid:
        csv_url = f"https://docs.google.com/spreadsheets/d/1p1aWr5N0GXdATgP9jM9pB_lL5HHsKEaQDU8zp0h7GMM/export?format=csv&gid={gid}"
        csv_urls.append(csv_url)
        print(f"Generated CSV export URL: {csv_url}")

# Sets to store link references
priority_link_references = set()  # Full words with delimiters
secondary_link_references = set()  # Smaller parts extracted from full words

# Function to split a value into parts based on delimiters
def split_value(value):
    parts = [value]  # Start with the full value
    for delimiter in delimiters:
        new_parts = []
        for part in parts:
            new_parts.extend(part.split(delimiter))
        parts = new_parts
    return [part.strip() for part in parts if part.strip()]

# Step 1: Download the CSV file
def download_csv(url):
    print(f"Downloading CSV from URL: {url}")
    response = requests.get(url)
    response.raise_for_status()  # Check for errors
    print("CSV downloaded successfully.")
    return StringIO(response.text)

# Function to sanitize a value for Markdown
def sanitize_value(value):
    # Remove leading/trailing whitespace
    value = value.strip()
    # Remove any trailing special characters (non-alphanumeric)
    value = re.sub(r"[^\w\s]+$", "", value)
    return value

def sanitize_filename(filename):
    # Apply sanitize_value to remove trailing special characters
    filename = sanitize_value(filename)
    # Replace problematic characters (except spaces) with underscores
    filename = re.sub(r'[\\/*?:"<>|]', '_', filename)
    return filename

def sanitize_link(link):
    # Apply sanitize_value to remove trailing special characters
    link = sanitize_value(link)
    return link

# Step 2: Create link reference sets
def create_link_references(csv_data):
    print("Creating link references...")
    try:
        reader = csv.DictReader(csv_data)
        if not reader.fieldnames:
            print("Warning: No headers found in the CSV file.")
            return
        
        # Use the first column header as the folder name (sanitized)
        folder_name = sanitize_value(reader.fieldnames[0])
        print(f"Using folder name: {folder_name}")
        
        # Add first column values to the priority_link_references set
        csv_data.seek(0)  # Reset the file pointer
        next(reader)  # Skip header
        for row in reader:
            if reader.fieldnames and row:  # Check if row and headers are not empty
                first_column_value = row[reader.fieldnames[0]]
                if first_column_value.strip():  # Skip empty cells
                    # Sanitize the value before adding to priority_link_references
                    sanitized_value = sanitize_value(first_column_value)
                    priority_link_references.add(f"{folder_name}/{sanitized_value}")
                    print(f"Added to priority link references: {folder_name}/{sanitized_value}")
    except Exception as e:
        print(f"Error processing CSV: {e}")

def process_csv(csv_data, sheet_index):
    print(f"Processing sheet {sheet_index + 1}...")
    try:
        reader = csv.reader(csv_data)
        headers = next(reader)  # Read the header row
        
        if not headers:
            print(f"Warning: No headers found in the CSV file for sheet {sheet_index + 1}.")
            return
        
        # Use the first column header as the folder name (sanitized)
        folder_name = sanitize_value(headers[0])
        print(f"Using folder name: {folder_name}")
        
        sheet_folder = os.path.join(vault_path, folder_name)
        
        # Create a folder for the sheet
        os.makedirs(sheet_folder, exist_ok=True)
        print(f"Created folder: {sheet_folder}")
        
        # Track header occurrences to handle duplicates
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
        
        # Loop through rows and create/overwrite Markdown files
        for row in reader:
            if not row:  # Skip empty rows
                print("Skipping empty row.")
                continue
            
            # Use the first column as the filename (sanitized)
            filename_value = sanitize_value(row[0].strip())
            if not filename_value:
                filename_value = "Untitled"
            
            # Create a valid filename (sanitize for Markdown)
            filename = sanitize_filename(filename_value)
            filename = f"{filename}.md"
            filepath = os.path.join(sheet_folder, filename)
            
            # Initialize the set of linked files for this file
            if filename not in link_dict:
                link_dict[filename] = set()
            
            # Write Markdown content (overwrite if file exists)
            with open(filepath, 'w', encoding='utf-8') as md_file:
                # Write front matter (YAML)
                md_file.write("---\n")
                for i, value in enumerate(row):
                    if value and value.strip():  # Only write non-empty fields
                        # Use the sanitized and numbered header
                        safe_key = sanitized_headers[i]
                        safe_value = sanitize_value(value)
                        md_file.write(f"{safe_key}: {safe_value}\n")
                md_file.write("---\n\n")
                
                # Write main content with links (only linked values)
                md_file.write("## Links\n")
                linked_values = set()  # Track linked values to avoid duplicates
                
                # Add links based on headers and values
                for i, value in enumerate(row):
                    if value and value.strip():  # Only process non-empty cells
                        # Use the sanitized and numbered header
                        sanitized_header = sanitized_headers[i]
                        sanitized_value = sanitize_value(value)
                        
                        # Skip if the link matches the current file's name
                        if sanitized_value == filename_value:
                            print(f"Skipping self-referencing link: [[{folder_name}/{sanitized_value}]]")
                            continue
                        
                        # Compare both the header and value to the link references
                        for reference in priority_link_references:
                            reference_filename = reference.split("/")[-1]  # Extract filename from reference
                            
                            # Check if the header matches the reference
                            if sanitized_header == reference_filename:
                                linked_values.add(f"[[{reference}]]")
                                link_dict[filename].add(reference)  # Add full reference (with folder)
                                print(f"Added link from header: [[{reference}]]")
                            
                            # Check if the value matches the reference
                            if sanitized_value == reference_filename:
                                linked_values.add(f"[[{reference}]]")
                                link_dict[filename].add(reference)  # Add full reference (with folder)
                                print(f"Added link from value: [[{reference}]]")
                
                # Write unique links
                for linked_value in sorted(linked_values):
                    md_file.write(f"- {linked_value}\n")
            
            print(f"Created/Overwritten: {filepath}")
    except Exception as e:
        print(f"Error processing sheet {sheet_index + 1}: {e}")

def update_reverse_links():
    print("Updating reverse links...")
    
    # Step 1: Add reverse links to link_dict
    for filename, links in link_dict.items():
        for link in links:
            # Extract the target filename from the link
            linked_filename = link.split("/")[-1] + ".md"
            if linked_filename in link_dict:
                # Add a reverse link to the target file
                reverse_link = f"{filename.replace('.md', '')}"
                link_dict[linked_filename].add(reverse_link)
                print(f"Added reverse link: [[{reverse_link}]] to {linked_filename}")
    
    # Step 2: Print link_dict for debugging
    print("\nlink_dict after adding reverse links:")
    for filename, links in link_dict.items():
        print(f"{filename}: {links}")
    
    # Step 3: Update each file with the reverse links
    for filename, links in link_dict.items():
        filepath = os.path.join(vault_path, filename)
        if os.path.exists(filepath):
            with open(filepath, 'r+', encoding='utf-8') as md_file:
                content = md_file.read()
                md_file.seek(0)  # Move to the beginning of the file
                
                # Check if the file already has a "## Links" section
                if "## Links" not in content:
                    # Add the "## Links" section at the end
                    md_file.seek(0, 2)  # Move to the end of the file
                    md_file.write("\n## Links\n")
                    for link in sorted(links):
                        md_file.write(f"- [[{link}]]\n")
                    print(f"Added '## Links' section to: {filepath}")
                else:
                    # If the section exists, append new links
                    md_file.seek(0)
                    lines = md_file.readlines()
                    md_file.seek(0)
                    md_file.truncate()  # Clear the file content
                    
                    links_section = False
                    for line in lines:
                        if line.strip() == "## Links":
                            links_section = True
                        md_file.write(line)
                    
                    if links_section:
                        # Append new links to the existing section
                        for link in sorted(links):
                            if f"[[{link}]]" not in content:  # Avoid duplicates
                                md_file.write(f"- [[{link}]]\n")
                        print(f"Appended links to: {filepath}")
        else:
            print(f"File not found: {filepath}")

def write_link_references():
    print("Writing link references to file...")
    link_reference_file = os.path.join(vault_path, "link_references.txt")
    with open(link_reference_file, 'w', encoding='utf-8') as f:
        f.write("Priority Link References:\n")
        for reference in sorted(priority_link_references):
            f.write(f"{reference}\n")
        f.write("\nSecondary Link References:\n")
        for reference in sorted(secondary_link_references):
            f.write(f"{reference}\n")
    print(f"Link references written to: {link_reference_file}")

# Main function
def main():
    print("Starting script...")
    
    # Step 0: Clean up the vault
    print("Step 0: Cleaning up the vault...")
    cleanup_vault(vault_path)
    
    # Step 1: Create link references
    print("Step 1: Creating link references...")
    for i, csv_url in enumerate(csv_urls):
        try:
            print(f"Processing sheet {i + 1}...")
            csv_data = download_csv(csv_url)
            create_link_references(csv_data)
        except Exception as e:
            print(f"Error downloading or processing CSV for sheet {i + 1}: {e}")
    
    # Step 2: Process each CSV
    print("Step 2: Processing each CSV...")
    for i, csv_url in enumerate(csv_urls):
        try:
            print(f"Processing sheet {i + 1}...")
            csv_data = download_csv(csv_url)
            process_csv(csv_data, i)
        except Exception as e:
            print(f"Error processing sheet {i + 1}: {e}")
    
    # Step 3: Update reverse links
    print("Step 3: Updating reverse links...")
    update_reverse_links()
    
    # Step 4: Write link references to file
    print("Step 4: Writing link references to file...")
    write_link_references()
    
    print("Script completed successfully.")

if __name__ == "__main__":
    main()
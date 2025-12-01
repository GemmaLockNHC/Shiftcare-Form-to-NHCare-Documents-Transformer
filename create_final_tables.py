#!/usr/bin/env python3

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.pdfgen import canvas
import csv
import os
import re
import io
import tempfile

# Try to import PDF parsing libraries
try:
    import pdfplumber  # text extraction
except Exception:
    pdfplumber = None
try:
    from pypdf import PdfReader  # form fields
except Exception:
    PdfReader = None

# Define custom colors
BLUE_COLOR = colors.HexColor('#316DB2')

def load_ndis_support_items():
    """Load NDIS support items from CSV file and return as a dictionary for lookup"""
    ndis_items = {}
    try:
        with open('outputs/other/NDIS Support Items - NDIS Support Items.csv', 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                # Use support item name as key for lookup
                item_name = row['Support Item Name'].strip()
                ndis_items[item_name] = {
                    'number': row['Support Item Number'].strip(),
                    'unit': row['Unit'].strip(),
                    'wa_price': row.get('WA', '').strip(),
                    'qld_price': row.get('QLD', '').strip()
                }
        print(f"DEBUG: Loaded {len(ndis_items)} NDIS support items from CSV")
        # Verify establishment fee item exists
        if "Establishment Fee For Personal Care/Participation" in ndis_items:
            est_fee = ndis_items["Establishment Fee For Personal Care/Participation"]
            print(f"DEBUG: Establishment fee item found - WA: {est_fee.get('wa_price')}, QLD: {est_fee.get('qld_price')}")
        else:
            print("DEBUG: WARNING - Establishment Fee For Personal Care/Participation not found in NDIS items")
    except FileNotFoundError:
        print("NDIS Support Items CSV file not found. Using placeholder data.")
    except Exception as e:
        print(f"Error loading NDIS support items: {e}")
    
    return ndis_items

def get_price_state(team_value):
    """
    Determine which state's prices to use based on the team.
    
    Args:
        team_value: The team name from 'Neighbourhood Care representative team'
    
    Returns:
        str: 'WA' or 'QLD'
    """
    if not team_value:
        return 'WA'  # Default to WA if team is not specified
    
    team_value_clean = team_value.strip().lower()
    
    # WA teams
    wa_teams = ['fremantle', 'belmont', 'metro-x', 'rockingham', 'wanneroo']
    if team_value_clean in wa_teams:
        return 'WA'
    
    # QLD teams
    qld_teams = ['brisbane', 'beaudesert', 'ipswich', 'gold coast']
    if team_value_clean in qld_teams:
        return 'QLD'
    
    # Default to WA if team doesn't match
    return 'WA'

def lookup_support_item(ndis_items, item_name):
    """Look up a support item by name and return its details"""
    if item_name in ndis_items:
        return ndis_items[item_name]
    else:
        # Try partial matching for common support items
        for key, value in ndis_items.items():
            if item_name.lower() in key.lower() or key.lower() in item_name.lower():
                return value
        # Return placeholder if not found
        return {
            'number': '[Not Found]',
            'unit': 'Hour',
            'wa_price': '$0.00',
            'qld_price': '$0.00'
        }

def get_establishment_fee(csv_data, ndis_items, team_value=None):
    """
    Calculate the establishment fee based on client status and support hours.
    
    The establishment fee is applied only if:
    - Client is new (isNewClient == "Yes")
    - Client is receiving 20+ hours of support (isReceiving20HoursSupport == "Yes" or similar)
    
    Args:
        csv_data: Dictionary containing form data
        ndis_items: Dictionary of NDIS support items loaded from CSV
        team_value: The team name from 'Neighbourhood Care representative team' (optional)
    
    Returns:
        str: Formatted establishment fee amount (e.g., "$702.30" or "$0.00")
    """
    # Check for new client status - try various possible field names
    # JavaScript code checked: submission.isNewClient == "Yes"
    is_new_client = False
    new_client_fields = [
        'isNewClient',  # Primary field name from JavaScript
        'Is this client new to Neighbourhood Care?',  # Actual PDF question text
        'Is this a new client?',
        'Is the client new?',
        'Is this client new?',
        'New client',
        'Is new client',
        'Client is new'
    ]
    
    for field in new_client_fields:
        value = csv_data.get(field, '').strip()
        # Clean up checkbox characters and other special characters
        value = value.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
        # Match JavaScript: exact "Yes" check (case-sensitive in JS, but we'll be flexible)
        if value == "Yes" or normalize_key(value) == 'yes':
            is_new_client = True
            break
    
    # Check for 20+ hours of support - try various possible field names
    # JavaScript code checked: submission.isReceiving20HoursSupport == "Yes"
    # Also had: submission.isReceiving20HoursOfSupport == "Yes"
    is_receiving_20_hours = False
    hours_support_fields = [
        'isReceiving20HoursSupport',  # Primary field name from JavaScript
        'isReceiving20HoursOfSupport',  # Alternative from JavaScript
        'Is Neighbourhood Care delivering 20 or more hours of support per month?',  # Actual PDF question text
        'Is the client receiving 20 or more hours of support?',
        'Is receiving 20 hours of support',
        'Is receiving 20+ hours of support',
        'Receiving 20 hours support',
        'Receiving 20+ hours',
        '20 hours support',
        '20 or more hours of support'
    ]
    
    for field in hours_support_fields:
        value = csv_data.get(field, '').strip()
        # Clean up checkbox characters and other special characters
        value = value.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
        # Match JavaScript: exact "Yes" check (case-sensitive in JS, but we'll be flexible)
        if value == "Yes" or normalize_key(value) == 'yes':
            is_receiving_20_hours = True
            break
    
    # If both conditions are met, get the establishment fee from NDIS support items
    if is_new_client and is_receiving_20_hours:
        # Look up "Establishment Fee For Personal Care/Participation" in NDIS items
        establishment_fee_item = lookup_support_item(ndis_items, "Establishment Fee For Personal Care/Participation")
        
        # Determine which state's price to use
        price_state = get_price_state(team_value)
        price_key = 'wa_price' if price_state == 'WA' else 'qld_price'
        price = establishment_fee_item.get(price_key, '$0.00')
        
        # Clean up the price string (remove any extra formatting)
        price = price.strip()
        
        # If price is empty or invalid, try to get it from the item directly
        if not price or price == '' or price == '$0.00':
            # Try the other price key as fallback
            other_key = 'qld_price' if price_key == 'wa_price' else 'wa_price'
            price = establishment_fee_item.get(other_key, '$0.00').strip()
        
        # Ensure price has $ prefix and proper formatting
        if price and price != '$0.00':
            if not price.startswith('$'):
                # If it's a number, add $ prefix
                try:
                    # Remove any commas and $ signs, then format
                    price_num = float(price.replace('$', '').replace(',', ''))
                    price = f"${price_num:.2f}"
                except (ValueError, AttributeError):
                    price = '$0.00'
            else:
                # Price already has $, just ensure proper formatting
                try:
                    price_num = float(price.replace('$', '').replace(',', ''))
                    price = f"${price_num:.2f}"
                except (ValueError, AttributeError):
                    price = '$0.00'
        
        # Debug output
        if price == '$0.00':
            print(f"DEBUG: Establishment fee calculation - is_new_client: {is_new_client}, is_receiving_20_hours: {is_receiving_20_hours}")
            print(f"DEBUG: Price state: {price_state}, Price key: {price_key}")
            print(f"DEBUG: Establishment fee item found: {establishment_fee_item}")
        
        return price
    else:
        # Debug output when conditions not met
        print(f"DEBUG: Establishment fee conditions not met - is_new_client: {is_new_client}, is_receiving_20_hours: {is_receiving_20_hours}")
        return '$0.00'

def load_active_users(team_value=None):
    """
    Load active users from CSV file and return as a dictionary for lookup.
    
    Args:
        team_value: The team name to determine which CSV file to use.
                   QLD teams (Beaudesert, Brisbane, Gold Coast, Ipswich) use Active_Users_1763520740.csv
                   Other teams use Active_Users_1761707021.csv
    
    Returns:
        Dictionary of active users keyed by name
    """
    active_users = {}
    
    # Determine which CSV file to use based on team
    qld_teams = ['beaudesert', 'brisbane', 'gold coast', 'ipswich']
    team_lower = team_value.strip().lower() if team_value else ''
    
    if team_lower in qld_teams:
        csv_filename = 'outputs/other/Active_Users_1763520740.csv'
        print(f"DEBUG: Using QLD active users CSV for team: {team_value}")
    else:
        csv_filename = 'outputs/other/Active_Users_1761707021.csv'
        print(f"DEBUG: Using default active users CSV for team: {team_value or 'unknown'}")
    
    try:
        with open(csv_filename, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                # Use name as key for lookup
                user_name = row['name'].strip()
                active_users[user_name] = {
                    'name': row['name'].strip(),
                    'mobile': row['mobile'].strip(),
                    'email': row['email'].strip(),
                    'team': (row.get('area') or row.get('role') or '').strip()
                }
        print(f"DEBUG: Loaded {len(active_users)} active users from {csv_filename}")
    except FileNotFoundError:
        print(f"Active Users CSV file not found: {csv_filename}. Using placeholder data.")
    except Exception as e:
        print(f"Error loading active users from {csv_filename}: {e}")
    
    return active_users

def lookup_user_data(active_users, respondent_name):
    """Look up user data by respondent name and return contact details"""
    if respondent_name in active_users:
        return active_users[respondent_name]
    else:
        # Try partial matching
        for key, value in active_users.items():
            if respondent_name.lower() in key.lower() or key.lower() in respondent_name.lower():
                return value
        # Return placeholder if not found
        return {
            'name': respondent_name,
            'mobile': '[Not Found]',
            'email': '[Not Found]'
        }

def normalize_key(key: str) -> str:
    """Normalize a key for comparison"""
    return str(key or "").strip().lower()

def extract_pdf_fields_pdfreader(pdf_path: str) -> dict:
    """Extract form fields from PDF using PdfReader"""
    if PdfReader is None:
        return {}
    try:
        reader = PdfReader(pdf_path)
        fields = {}
        if reader.get_fields():
            for name, field in reader.get_fields().items():
                value = field.get("/V")
                if value is None:
                    continue
                fields[name] = str(value)
        return fields
    except Exception:
        return {}

def extract_pdf_text_pdfplumber(pdf_path: str) -> str:
    """Extract all text from PDF using pdfplumber"""
    if pdfplumber is None:
        return ""
    try:
        text_parts = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text_parts.append(page.extract_text() or "")
        return "\n".join(text_parts)
    except Exception:
        return ""

def parse_pdf_to_data(pdf_path: str) -> dict:
    """Parse PDF and extract data, mapping to CSV field names"""
    data = {}
    
    # Debug flag - set to True to see what's being extracted
    DEBUG = False
    
    # First try to extract form fields
    fields = extract_pdf_fields_pdfreader(pdf_path)
    if fields:
        # Debug: print all field names found (commented out for production)
        # print("PDF Form Fields found:", list(fields.keys()))
        
        # Map form fields to CSV field names
        # The PDF form field NAMES are simple like "First name", "Home address", etc.
        # The VALUES in those fields are the actual data (might look like "First name (Details of the Client)" but treat as real data)
        def find_in_fields(*candidates):
            # Try all candidates - match simple field names
            for cand in candidates:
                cand_norm = normalize_key(cand)
                for key, val in fields.items():
                    key_norm = normalize_key(key)
                    value = str(val).strip()
                    if not value:
                        continue
                    
                    # Exact match
                    if cand_norm == key_norm:
                        return value
                    
                    # Substring match - candidate is in the key
                    if cand_norm in key_norm:
                        return value
            return ""
        
        # Details of the Client section - match simple field names like "First name"
        data['First name (Details of the Client)'] = find_in_fields("first name", "firstname")
        data['Middle name (Details of the Client)'] = find_in_fields("middle name", "middlename")
        data['Surname (Details of the Client)'] = find_in_fields("surname", "family name", "last name", "lastname")
        data['NDIS number (Details of the Client)'] = find_in_fields("ndis number", "ndis")
        data['Date of birth (Details of the Client)'] = find_in_fields("date of birth", "dob", "birth date")
        data['Gender (Details of the Client)'] = find_in_fields("gender")
        
        # Contact Details of the Client section - match simple field names like "Home address"
        data['Home address (Contact Details of the Client)'] = find_in_fields("home address", "address")
        data['Home phone (Contact Details of the Client)'] = find_in_fields("home phone", "homephone")
        data['Work phone (Contact Details of the Client)'] = find_in_fields("work phone", "workphone")
        data['Mobile phone (Contact Details of the Client)'] = find_in_fields("mobile phone", "mobile", "mobilephone")
        data['Email address (Contact Details of the Client)'] = find_in_fields("email address", "email")
        
        # Emergency contact - match simple field names, but need to distinguish from client fields
        # Try emergency-specific names first, then fallback to simple names
        data['First name (Emergency contact)'] = find_in_fields("first name (emergency contact)", "emergency contact first name", "emergency first name")
        if not data['First name (Emergency contact)']:
            # If we found a "first name" field, check if there's an emergency-specific one
            # For now, we'll rely on text extraction to get the right one
            pass
        data['Surname (Emergency contact)'] = find_in_fields("surname (emergency contact)", "emergency contact surname", "emergency contact last name", "emergency surname", "emergency last name")
        data['Is the primary carer also the emergency contact for the participant?'] = find_in_fields("primary carer also emergency contact", "is primary carer emergency contact")
        data['Home phone (Emergency contact)'] = find_in_fields("home phone (emergency contact)", "emergency contact home phone", "emergency home phone")
        data['Mobile phone (Emergency contact)'] = find_in_fields("mobile phone (emergency contact)", "emergency contact mobile phone", "emergency mobile phone", "emergency contact mobile")
        data['Work phone (Emergency contact)'] = find_in_fields("work phone (emergency contact)", "emergency contact work phone", "emergency work phone")
        data['Relationship to client (Emergency contact)'] = find_in_fields("relationship to client (emergency contact)", "emergency contact relationship", "emergency relationship")
        
        # Extract Person Signing the Agreement fields
        person_signing = find_in_fields("person signing the agreement", "who is signing", "signatory")
        # Clean up checkbox characters
        if person_signing:
            person_signing = person_signing.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
        data['Person signing the agreement'] = person_signing
        data['First name (Person Signing the Agreement)'] = find_in_fields("first name (person signing the agreement)", "first name (person signing", "person signing first name", "signatory first name")
        data['Surname (Person Signing the Agreement)'] = find_in_fields("surname (person signing the agreement)", "surname (person signing", "person signing surname", "person signing last name", "signatory surname", "signatory last name")
        data['Relationship to client (Person Signing the Agreement)'] = find_in_fields("relationship to client (person signing the agreement)", "relationship to client (person signing", "person signing relationship", "signatory relationship")
        data['Home address (Person Signing the Agreement)'] = find_in_fields("home address (person signing the agreement)", "home address (person signing", "person signing address", "signatory address")
        data['Home phone (Person Signing the Agreement)'] = find_in_fields("home phone (person signing the agreement)", "home phone (person signing", "person signing home phone", "signatory home phone")
        data['Mobile phone (Person Signing the Agreement)'] = find_in_fields("mobile phone (person signing the agreement)", "mobile phone (person signing", "person signing mobile", "signatory mobile")
        data['Email address (Person Signing the Agreement)'] = find_in_fields("email address (person signing the agreement)", "email address (person signing", "person signing email", "signatory email")
        
        # Extract Primary carer fields
        data['First name (Primary carer)'] = find_in_fields("first name (primary carer)", "first name (primary carer", "primary carer first name")
        data['Surname (Primary carer)'] = find_in_fields("surname (primary carer)", "surname (primary carer", "primary carer surname", "primary carer last name")
        data['Relationship to client (Primary carer)'] = find_in_fields("relationship to client (primary carer)", "relationship to client (primary carer", "primary carer relationship")
        data['Home address (Primary carer)'] = find_in_fields("home address (primary carer)", "home address (primary carer", "primary carer address")
        data['Home phone (Primary carer)'] = find_in_fields("home phone (primary carer)", "home phone (primary carer", "primary carer home phone")
        data['Mobile phone (Primary carer)'] = find_in_fields("mobile phone (primary carer)", "mobile phone (primary carer", "primary carer mobile")
        data['Email address (Primary carer)'] = find_in_fields("email address (primary carer)", "email address (primary carer", "primary carer email")
    
    # Always try text extraction as well to fill in any missing fields
    # This ensures we get all fields even if form field extraction missed some
    text = extract_pdf_text_pdfplumber(pdf_path)
    if text:
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        
        # Identify section boundaries
        section_starts = []
        for i, line in enumerate(lines):
            line_lower = normalize_key(line)
            line_clean = line.strip()
            
            # Only match actual section headers - they should be standalone lines without parentheses
            # Section headers are typically short and don't contain field values
            is_section_header = (
                len(line_clean) < 50 and  # Section headers are short
                '(' not in line_clean and  # Section headers don't have parentheses (values do)
                ':' not in line_clean      # Section headers don't have colons
            )
            
            if is_section_header:
                if "details of the client" in line_lower and "contact" not in line_lower:
                    section_starts.append(("details", i))
                elif "contact details of the client" in line_lower:
                    section_starts.append(("contact", i))
                elif "primary carer" in line_lower and "emergency" not in line_lower:
                    section_starts.append(("primary_carer", i))
                elif "emergency contact" in line_lower:
                    section_starts.append(("emergency", i))
                elif any(x in line_lower for x in ["needs of the client", "ndis information", "support items", "formal supports", 
                                                   "important people", "home life", "health information", 
                                           "care requirements", "behaviour requirements", "other information", "consents"]):
                    # End of relevant sections - but only break if we've found primary_carer and emergency sections
                    if any(sec[0] == "primary_carer" for sec in section_starts) and any(sec[0] == "emergency" for sec in section_starts):
                        break
        
        # Helper function to find value in a specific section - SIMPLIFIED
        def find_value_in_section(label_patterns, section_type):
            """Find value only in the specified section - just get the next line after the label"""
            # Find the relevant section
            section_start = None
            section_end = None
            
            for sec_type, start_idx in section_starts:
                if sec_type == section_type:
                    section_start = start_idx
                    # Find end of this section (start of next section or end of lines)
                    for next_sec_type, next_start_idx in section_starts:
                        if next_start_idx > start_idx:
                            section_end = next_start_idx
                            break
                    if section_end is None:
                        section_end = len(lines) if lines else 0
                    break
            
            if section_start is None or section_end is None:
                return ""
            
            # Ensure both are integers
            section_start = int(section_start) if section_start is not None else 0
            section_end = int(section_end) if section_end is not None else len(lines) if lines else 0
            
            # Only search within this section
            for i in range(section_start, section_end):
                line = lines[i]
                line_lower = normalize_key(line)
                
                for pattern in label_patterns:
                    pattern_lower = normalize_key(pattern)
                    
                    # Simple match - check if pattern matches the line
                    line_clean = line_lower.replace("(details of the client)", "").replace("(contact details of the client)", "").strip()
                    pattern_clean = pattern_lower.replace("(details of the client)", "").replace("(contact details of the client)", "").strip()
                    
                    matches = (
                        pattern_lower == line_lower or
                        pattern_clean == line_clean or
                        line_lower.startswith(pattern_lower)
                    )
                    
                    if matches:
                        # Look for value on same line after colon (if present)
                        if ':' in line:
                            parts = line.split(':', 1)
                            if len(parts) > 1 and parts[1].strip():
                                return parts[1].strip()
                        
                        # Just get the next non-empty line - that's the value
                        # But skip if it's clearly another field label
                        field_labels = ['first name', 'middle name', 'surname', 'ndis number', 'date of birth', 'gender',
                                       'home address', 'home phone', 'work phone', 'mobile phone', 'email address',
                                       'preferred name', 'key code', 'postal address', 'preferred method of contact',
                                       'relationship to client']
                        section_end_safe = section_end if section_end is not None and isinstance(section_end, int) else (len(lines) if lines else i + 5)
                        for j in range(i + 1, min(i + 5, section_end_safe)):
                            next_line = lines[j].strip()
                            if not next_line or next_line in ['•', '●', '○', '☐', '☑', '✓']:
                                continue
                            
                            # Skip if it's another field label
                            next_line_lower = normalize_key(next_line)
                            is_field_label = False
                            for fl in field_labels:
                                if fl == next_line_lower or (fl in next_line_lower and len(next_line) < 50 and '(' not in next_line):
                                    is_field_label = True
                                    break
                            
                            # Skip instruction text
                            if len(next_line) > 80 or any(x in next_line_lower for x in ['write', 'below', 'same as', 'if their']):
                                continue
                            
                            if not is_field_label:
                                return next_line
            
            return ""
        
        # Helper function for fields that aren't in specific sections - SIMPLIFIED
        def find_value_after_label(label_patterns, start_idx=0):
            for i in range(start_idx, len(lines) if lines else 0):
                line_lower = normalize_key(lines[i])
                for pattern in label_patterns:
                    pattern_lower = normalize_key(pattern)
                    # Match if pattern is in the line (but not if the line IS the pattern - that's the label)
                    if pattern_lower in line_lower:
                        # Look for value on same line after colon
                        if ':' in lines[i]:
                            parts = lines[i].split(':', 1)
                            if len(parts) > 1 and parts[1].strip():
                                return parts[1].strip()
                        # Skip the label line itself and get the next non-empty line - that's the value
                        for j in range(i + 1, min(i + 3, len(lines) if lines else i + 3)):
                            next_line = lines[j].strip()
                            if next_line and next_line not in ['•', '●', '○', '☐', '☑', '✓']:
                                # Make sure we're not returning the label itself
                                next_line_lower = normalize_key(next_line)
                                if next_line_lower != pattern_lower:
                                    return next_line
            return ""
    
        # Extract data using section-aware text parsing - only fill in missing fields
        if not data.get('First name (Details of the Client)'):
            data['First name (Details of the Client)'] = find_value_in_section(['First name', 'First name (Details of the Client)'], "details")
        if not data.get('Middle name (Details of the Client)'):
            data['Middle name (Details of the Client)'] = find_value_in_section(['Middle name', 'Middle name (Details of the Client)'], "details")
        if not data.get('Surname (Details of the Client)'):
            data['Surname (Details of the Client)'] = find_value_in_section(['Surname', 'Surname (Details of the Client)', 'Family name', 'Last name'], "details")
        if not data.get('NDIS number (Details of the Client)'):
            data['NDIS number (Details of the Client)'] = find_value_in_section(['NDIS number', 'NDIS number (Details of the Client)'], "details")
        if not data.get('Date of birth (Details of the Client)'):
            data['Date of birth (Details of the Client)'] = find_value_in_section(['Date of birth', 'Date of birth (Details of the Client)', 'DOB'], "details")
        if not data.get('Gender (Details of the Client)'):
            data['Gender (Details of the Client)'] = find_value_in_section(['Gender', 'Gender (Details of the Client)'], "details")
        if not data.get('Home address (Contact Details of the Client)'):
            data['Home address (Contact Details of the Client)'] = find_value_in_section(['Home address', 'Home address (Contact Details of the Client)', 'Address'], "contact")
        if not data.get('Home phone (Contact Details of the Client)'):
            data['Home phone (Contact Details of the Client)'] = find_value_in_section(['Home phone', 'Home phone (Contact Details of the Client)'], "contact")
        if not data.get('Work phone (Contact Details of the Client)'):
            data['Work phone (Contact Details of the Client)'] = find_value_in_section(['Work phone', 'Work phone (Contact Details of the Client)'], "contact")
        if not data.get('Mobile phone (Contact Details of the Client)'):
            data['Mobile phone (Contact Details of the Client)'] = find_value_in_section(['Mobile phone', 'Mobile phone (Contact Details of the Client)'], "contact")
        if not data.get('Email address (Contact Details of the Client)'):
            data['Email address (Contact Details of the Client)'] = find_value_in_section(['Email address', 'Email address (Contact Details of the Client)', 'Email'], "contact")
    
        # Extract emergency contact fields - try emergency section first, then fallback to general search
        if not data.get('First name (Emergency contact)'):
            emergency_first = find_value_in_section(['First name'], "emergency")
            if emergency_first:
                data['First name (Emergency contact)'] = emergency_first
            else:
                data['First name (Emergency contact)'] = find_value_after_label(['First name (Emergency contact)'])
        if not data.get('Surname (Emergency contact)'):
            emergency_surname = find_value_in_section(['Surname'], "emergency")
            if emergency_surname:
                data['Surname (Emergency contact)'] = emergency_surname
            else:
                data['Surname (Emergency contact)'] = find_value_after_label(['Surname (Emergency contact)'])
        if not data.get('Is the primary carer also the emergency contact for the participant?'):
            data['Is the primary carer also the emergency contact for the participant?'] = find_value_after_label(['Is the primary carer also the emergency contact'])
        # Extract emergency contact phone and relationship fields
        if not data.get('Home phone (Emergency contact)'):
            data['Home phone (Emergency contact)'] = find_value_in_section(['Home phone'], "emergency")
        if not data.get('Mobile phone (Emergency contact)'):
            data['Mobile phone (Emergency contact)'] = find_value_in_section(['Mobile phone'], "emergency")
        if not data.get('Work phone (Emergency contact)'):
            data['Work phone (Emergency contact)'] = find_value_in_section(['Work phone'], "emergency")
        if not data.get('Relationship to client (Emergency contact)'):
            data['Relationship to client (Emergency contact)'] = find_value_in_section(['Relationship to client', 'Relationship'], "emergency")
    
    # Extract other fields that might be in the PDF
        if not data.get('Preferred method of contact'):
            data['Preferred method of contact'] = find_value_after_label(['Preferred method of contact', 'Preferred contact method'])
        if not data.get('Total core budget to allocate to Neighbourhood Care'):
            data['Total core budget to allocate to Neighbourhood Care'] = find_value_after_label(['Total core budget', 'core budget'])
        if not data.get('Total capacity building budget to allocate to Neighbourhood Care'):
            data['Total capacity building budget to allocate to Neighbourhood Care'] = find_value_after_label(['Total capacity building budget', 'capacity building budget'])
        if not data.get('Plan start date'):
            data['Plan start date'] = find_value_after_label(['Plan start date', 'Plan start'])
        if not data.get('Plan end date'):
            data['Plan end date'] = find_value_after_label(['Plan end date', 'Plan end'])
        if not data.get('Service start date'):
            data['Service start date'] = find_value_after_label(['Service start date', 'Service start'])
        if not data.get('Service end date'):
            data['Service end date'] = find_value_after_label(['Service end date', 'Service end'])
        # Always try text extraction for Person signing the agreement (form fields might return the label)
        person_signing_text = find_value_after_label(['Person signing the agreement', 'Who is signing'])
        if person_signing_text and person_signing_text.lower() != 'person signing the agreement':
            # Clean up checkbox characters
            person_signing_text = person_signing_text.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
            if person_signing_text:
                data['Person signing the agreement'] = person_signing_text
        if not data.get('First name (Person Signing the Agreement)'):
            data['First name (Person Signing the Agreement)'] = find_value_after_label(['First name (Person Signing the Agreement)'])
        if not data.get('Surname (Person Signing the Agreement)'):
            data['Surname (Person Signing the Agreement)'] = find_value_after_label(['Surname (Person Signing the Agreement)'])
        if not data.get('Relationship to client (Person Signing the Agreement)'):
            data['Relationship to client (Person Signing the Agreement)'] = find_value_after_label(['Relationship to client (Person Signing the Agreement)', 'Relationship'])
        if not data.get('Home address (Person Signing the Agreement)'):
            data['Home address (Person Signing the Agreement)'] = find_value_after_label(['Home address (Person Signing the Agreement)'])
        if not data.get('First name (Primary carer)'):
            # Look for "First name" label in Primary carer section
            data['First name (Primary carer)'] = find_value_in_section(['First name'], "primary_carer")
        if not data.get('Surname (Primary carer)'):
            # Look for "Surname" label in Primary carer section  
            data['Surname (Primary carer)'] = find_value_in_section(['Surname'], "primary_carer")
        if not data.get('Relationship to client (Primary carer)'):
            data['Relationship to client (Primary carer)'] = find_value_after_label(['Relationship to client (Primary carer)'])
        if not data.get('Home address (Primary carer)'):
            data['Home address (Primary carer)'] = find_value_after_label(['Home address (Primary carer)'])
        if not data.get('Plan management type'):
            data['Plan management type'] = find_value_after_label(['Plan management type', 'Plan management'])
        if not data.get('Plan manager name'):
            data['Plan manager name'] = find_value_after_label(['Plan manager name'])
        
        # Extract support items from Support Items Required section
        for i in range(1, 20):
            key = f'Support item ({i}) (Support Items Required)'
            if not data.get(key):
                # Look for "Support item (X)" label and get the value
                label_pattern = f'Support item ({i})'
                value = find_value_after_label([label_pattern])
                if value:
                    data[key] = value
        if not data.get('Plan manager postal address'):
            data['Plan manager postal address'] = find_value_after_label(['Plan manager postal address', 'Plan manager address'])
        if not data.get('Plan manager phone number'):
            data['Plan manager phone number'] = find_value_after_label(['Plan manager phone', 'Plan manager phone number'])
        if not data.get('Plan manager email address'):
            data['Plan manager email address'] = find_value_after_label(['Plan manager email'])
        if not data.get('Respondent'):
            data['Respondent'] = find_value_after_label(['Respondent', 'Neighbourhood Care representative'])
        if not data.get('Neighbourhood Care representative team'):
            data['Neighbourhood Care representative team'] = find_value_after_label(['Neighbourhood Care representative team', 'Team'])
        
        # Extract establishment fee related fields
        if not data.get('Is this client new to Neighbourhood Care?'):
            data['Is this client new to Neighbourhood Care?'] = find_value_after_label(['Is this client new to Neighbourhood Care?', 'Is this client new', 'Is this a new client'])
        if not data.get('Is Neighbourhood Care delivering 20 or more hours of support per month?'):
            data['Is Neighbourhood Care delivering 20 or more hours of support per month?'] = find_value_after_label(['Is Neighbourhood Care delivering 20 or more hours of support per month?', 'Is Neighbourhood Care delivering 20', '20 or more hours of support'])
    
        # Extract consent responses - look for Yes/No patterns
        consent_labels = [
            'I agree to receive services from Neighbourhood Care.',
            'I consent for Neighbourhood Care to create an NDIS portal service booking',
            'I understand that if at any time I (The Participant) require emergency medical assistance',
            'I agree that Neighbourhood Care staff may administer simple first aid',
            'I consent for Neighbourhood Care to discuss relevant information',
            'I agree not to smoke inside the home',
            'I understand that an Emergency Response Plan will be developed',
            'I consent for Neighbourhood Care for I (The Participant) to be photographed',
            'I give authority for my details or information to be shared'
        ]
        
        for consent_label in consent_labels:
            if consent_label not in data:
                # Look for the consent text and find Yes/No after it
                for i, line in enumerate(lines):
                    if normalize_key(consent_label.split('.')[0]) in normalize_key(line):
                        # Look for Yes/No in nearby lines
                        for j in range(max(0, i-2), min(len(lines) if lines else i+5, i+5)):
                            if normalize_key(lines[j]) in ['yes', 'no']:
                                data[consent_label] = lines[j]
                        break
    
        # Debug output
        if DEBUG:
            print("\n=== Extracted Data ===")
            for key, value in data.items():
                if value:
                    print(f"{key}: {value}")
            print("=====================\n")
    
    return data

def get_preferred_contact_details(csv_data):
    """Get contact details based on preferred method of contact"""
    preferred_method = csv_data.get('Preferred method of contact', '').lower()
    
    if 'home phone' in preferred_method:
        return csv_data.get('Home phone (Contact Details of the Client)', 'Home phone (Contact Details of the Client)')
    elif 'mobile' in preferred_method:
        return csv_data.get('Mobile phone (Contact Details of the Client)', 'Mobile phone (Contact Details of the Client)')
    elif 'work phone' in preferred_method:
        return csv_data.get('Work phone (Contact Details of the Client)', 'Work phone (Contact Details of the Client)')
    elif 'email' in preferred_method:
        return csv_data.get('Email address (Contact Details of the Client)', 'Email address (Contact Details of the Client)')
    else:
        # Default to home phone if no clear preference
        return csv_data.get('Home phone (Contact Details of the Client)', 'Home phone (Contact Details of the Client)')

# Signature extraction removed to prevent timeouts
def _extract_signatures_from_pdf_removed(source_pdf_path):
    """
    Extract signature images from the source PDF.
    Returns a dictionary with signature images (file paths).
    Signatures can be stored as:
    1. Form field signature values
    2. Embedded images in the PDF
    3. Annotations
    
    Uses multiple methods with fallbacks:
    - PyMuPDF (if available) - best for image extraction
    - pypdf - for form field signatures and embedded images
    - pdfplumber - for text-based identification
    
    This function is designed to fail gracefully - if extraction fails,
    it returns an empty dict and the PDF generation continues without signatures.
    
    OPTIMIZED: Only processes last 5 pages (where signatures usually are) to avoid timeouts.
    """
    import signal
    import time
    
    signatures = {}
    if not source_pdf_path or not os.path.exists(source_pdf_path):
        print(f"Signature extraction: Source PDF not found: {source_pdf_path}")
        return signatures
    
    print(f"Signature extraction: Attempting to extract from {source_pdf_path}")
    
    # Set a maximum processing time (30 seconds)
    start_time = time.time()
    MAX_PROCESSING_TIME = 30
    MAX_PAGES_TO_PROCESS = 5  # Only check last 5 pages (signatures are usually at the end)
    
    try:
        # Method 1: Try PyMuPDF first (best method, but optional)
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(source_pdf_path)
            image_list = []
            
            total_pages = len(doc)
            print(f"Signature extraction: Using PyMuPDF, found {total_pages} pages")
            
            # Only process last MAX_PAGES_TO_PROCESS pages (signatures are usually at the end)
            start_page = max(0, total_pages - MAX_PAGES_TO_PROCESS)
            print(f"Signature extraction: Processing pages {start_page + 1} to {total_pages} (last {MAX_PAGES_TO_PROCESS} pages)")
            
            for page_num in range(start_page, total_pages):
                # Check timeout
                if time.time() - start_time > MAX_PROCESSING_TIME:
                    print("Signature extraction: Timeout reached during PyMuPDF processing")
                    break
                
                page = doc[page_num]
                # Get all images on the page
                page_images = page.get_images()
                print(f"Signature extraction: Page {page_num + 1} has {len(page_images)} images")
                # Limit images per page
                image_list.extend(page_images[:10])  # Max 10 images per page
            
            print(f"Signature extraction: Total images found: {len(image_list)}")
            
            # Extract images that might be signatures
            # Look for images at the bottom of pages (where signatures usually are)
            # Limit to first 20 images total to avoid memory issues
            for img_index, img in enumerate(image_list[:20]):
                # Check timeout
                if time.time() - start_time > MAX_PROCESSING_TIME:
                    print("Signature extraction: Timeout reached during image extraction")
                    break
                
                # Stop if we already have 2 signatures
                if len(signatures) >= 2:
                    break
                
                try:
                    # Get image data
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    # Save to temporary file
                    with tempfile.NamedTemporaryFile(delete=False, suffix=f'.{image_ext}') as tmp_file:
                        tmp_file.write(image_bytes)
                        tmp_path = tmp_file.name
                    
                    print(f"Signature extraction: Extracted image {img_index + 1} ({len(image_bytes)} bytes) to {tmp_path}")
                    
                    # Take first 2 images as potential signatures
                    if img_index < 2:
                        key = 'signatory' if img_index == 0 else 'nc_representative'
                        signatures[key] = tmp_path
                        print(f"Signature extraction: Assigned image {img_index + 1} as {key}")
                except Exception as e:
                    print(f"Error extracting image {img_index}: {e}")
            
            doc.close()
            if signatures:
                print(f"Signature extraction: Successfully extracted {len(signatures)} signatures using PyMuPDF")
                return signatures
            else:
                print("Signature extraction: PyMuPDF found images but none were assigned as signatures")
        except ImportError:
            # PyMuPDF not available, continue to other methods
            print("Signature extraction: PyMuPDF not available, trying pypdf methods")
            pass
        except Exception as e:
            print(f"Error using PyMuPDF: {e}")
        
        # Method 2: Try to extract from signature form fields using pypdf
        if PdfReader is not None:
            try:
                reader = PdfReader(source_pdf_path)
                
                if reader.get_fields():
                    for field_name, field in reader.get_fields().items():
                        field_name_lower = (field_name or "").lower()
                        
                        # Check if this is a signature field
                        if 'signature' in field_name_lower or 'sign' in field_name_lower:
                            try:
                                # Try to extract signature image from appearance stream
                                if hasattr(field, 'get'):
                                    ap = field.get('/AP')
                                    if ap:
                                        normal_ap = ap.get('/N')
                                        if normal_ap:
                                            # Try to get image from appearance stream
                                            try:
                                                # Access the stream object
                                                if hasattr(normal_ap, 'get_data'):
                                                    stream_data = normal_ap.get_data()
                                                    # Try to extract image from PDF stream
                                                    # This is complex - would need PDF stream parsing
                                                    pass
                                                elif hasattr(normal_ap, 'get_object'):
                                                    # Try to get the object and extract image
                                                    obj = normal_ap.get_object()
                                                    if obj and hasattr(obj, 'get_data'):
                                                        stream_data = obj.get_data()
                                                        # Would need to parse PDF stream to extract image
                                                        pass
                                            except Exception as e:
                                                print(f"Error extracting from appearance stream: {e}")
                            except Exception as e:
                                print(f"Error extracting signature from field {field_name}: {e}")
            except Exception as e:
                print(f"Error reading PDF fields: {e}")
        
        # Method 3: Try using pdfplumber to extract images (better for FlateDecode images)
        if not signatures:
            try:
                # Check timeout
                if time.time() - start_time > MAX_PROCESSING_TIME:
                    print("Signature extraction: Timeout reached, skipping pdfplumber method")
                    return signatures
                
                print("Signature extraction: Trying pdfplumber image extraction method")
                import pdfplumber
                with pdfplumber.open(source_pdf_path) as pdf:
                    image_count = 0
                    total_images_found = 0
                    total_pages = len(pdf.pages)
                    
                    # Only process last MAX_PAGES_TO_PROCESS pages (signatures are usually at the end)
                    start_page = max(0, total_pages - MAX_PAGES_TO_PROCESS)
                    print(f"Signature extraction: Processing pages {start_page + 1} to {total_pages} (last {MAX_PAGES_TO_PROCESS} pages)")
                    
                    for page_num in range(start_page, total_pages):
                        # Check timeout before processing each page
                        if time.time() - start_time > MAX_PROCESSING_TIME:
                            print("Signature extraction: Timeout reached during page processing")
                            break
                        
                        page = pdf.pages[page_num]
                        images = page.images
                        if images:
                            print(f"Signature extraction: Page {page_num + 1} has {len(images)} images")
                            total_images_found += len(images)
                            
                            # Limit number of images processed per page
                            max_images_per_page = 10
                            images_to_process = images[:max_images_per_page]
                            
                            for img in images_to_process:
                                # Check timeout before processing each image
                                if time.time() - start_time > MAX_PROCESSING_TIME:
                                    print("Signature extraction: Timeout reached during image processing")
                                    break
                                
                                # Stop if we already have 2 signatures
                                if image_count >= 2:
                                    break
                                
                                try:
                                    # Check if this looks like a signature (usually at bottom of page, specific size)
                                    # Signatures are typically wider than tall, and positioned near bottom
                                    img_height = img.get('height', 0) or 0
                                    img_width = img.get('width', 0) or 0
                                    y_position = img.get('y0', 0) or 0
                                    page_height = getattr(page, 'height', None) or 800  # Default to 800 if None
                                    
                                    # Signatures are usually at bottom 20% of page and have reasonable dimensions
                                    is_likely_signature = (
                                        img_height and img_height > 20 and img_height < 200 and  # Reasonable signature height
                                        img_width and img_width > 100 and  # Signatures are usually wide
                                        page_height and y_position < page_height * 0.3  # Near bottom of page (y0 is from bottom)
                                    )
                                    
                                    if is_likely_signature or image_count < 2:  # Take first 2 images if we can't determine
                                        stream = img.get('stream')
                                        if stream:
                                            try:
                                                # Try to get raw image data
                                                if hasattr(stream, 'get_data'):
                                                    image_data = stream.get_data()
                                                elif hasattr(stream, '_data'):
                                                    image_data = stream._data
                                                else:
                                                    # Try to extract from the stream object
                                                    image_data = None
                                                    try:
                                                        # For FlateDecode, we need to decompress
                                                        import zlib
                                                        if hasattr(stream, 'raw_bytes'):
                                                            raw_bytes = stream.raw_bytes
                                                            # Try to decompress if it's compressed
                                                            try:
                                                                image_data = zlib.decompress(raw_bytes)
                                                            except:
                                                                image_data = raw_bytes
                                                    except Exception as e:
                                                        print(f"Error extracting stream data: {e}")
                                                
                                                if image_data and len(image_data) > 100:
                                                    # Determine file extension based on filter
                                                    suffix = '.png'  # Default for FlateDecode
                                                    if hasattr(stream, 'get'):
                                                        filter_type = stream.get('/Filter', '')
                                                        if '/DCTDecode' in (filter_type if isinstance(filter_type, list) else [filter_type]):
                                                            suffix = '.jpg'
                                                    
                                                    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                                                        tmp_file.write(image_data)
                                                        tmp_path = tmp_file.name
                                                    
                                                    print(f"Signature extraction: Saved image from page {page_num + 1} ({len(image_data)} bytes) to {tmp_path}")
                                                    
                                                    if image_count < 2:
                                                        key = 'signatory' if image_count == 0 else 'nc_representative'
                                                        signatures[key] = tmp_path
                                                        image_count += 1
                                                        print(f"Signature extraction: Assigned image as {key}")
                                            except Exception as e:
                                                print(f"Error extracting image data: {e}")
                                                import traceback
                                                traceback.print_exc()
                                except Exception as e:
                                    print(f"Error processing image: {e}")
                    
                    print(f"Signature extraction: pdfplumber found {total_images_found} total images, extracted {len(signatures)} as signatures")
            except Exception as e:
                print(f"Error extracting images with pdfplumber: {e}")
                import traceback
                traceback.print_exc()
        
        # Method 4: Try using pypdf to extract images from pages (fallback)
        if not signatures and PdfReader is not None:
            try:
                print("Signature extraction: Trying pypdf image extraction method (fallback)")
                reader = PdfReader(source_pdf_path)
                image_count = 0
                total_images_found = 0
                
                total_pages = len(reader.pages)
                # Only process last MAX_PAGES_TO_PROCESS pages
                start_page = max(0, total_pages - MAX_PAGES_TO_PROCESS)
                print(f"Signature extraction: Processing pages {start_page + 1} to {total_pages} (last {MAX_PAGES_TO_PROCESS} pages)")
                
                for page_num in range(start_page, total_pages):
                    # Check timeout
                    if time.time() - start_time > MAX_PROCESSING_TIME:
                        print("Signature extraction: Timeout reached during pypdf processing")
                        break
                    
                    # Stop if we already have 2 signatures
                    if len(signatures) >= 2:
                        break
                    
                    try:
                        page = reader.pages[page_num]
                        resources = page.get('/Resources', {})
                        if resources and '/XObject' in resources:
                            xobjects = resources['/XObject']
                            if hasattr(xobjects, 'get_object'):
                                xobjects = xobjects.get_object()
                            
                            if isinstance(xobjects, dict):
                                for obj_name, obj in xobjects.items():
                                    try:
                                        if hasattr(obj, 'get') and obj.get('/Subtype') == '/Image':
                                            total_images_found += 1
                                            print(f"Signature extraction: Found image object {obj_name} on page {page_num + 1}")
                                            
                                            # Extract image data - try multiple filter types
                                            filter_type = obj.get('/Filter', '')
                                            is_supported = False
                                            
                                            # Check for JPEG (DCTDecode)
                                            if filter_type == '/DCTDecode' or (isinstance(filter_type, list) and '/DCTDecode' in filter_type):
                                                is_supported = True
                                                suffix = '.jpg'
                                            # Check for PNG (FlateDecode) - this is what signatures use!
                                            elif filter_type == '/FlateDecode' or (isinstance(filter_type, list) and '/FlateDecode' in filter_type):
                                                is_supported = True
                                                suffix = '.png'
                                            # Check for other common formats
                                            elif '/CCITTFaxDecode' in (filter_type if isinstance(filter_type, list) else [filter_type]):
                                                is_supported = True
                                                suffix = '.tiff'
                                            
                                            if is_supported:
                                                try:
                                                    image_data = obj.get_data()
                                                    if image_data and len(image_data) > 100:  # Only save if it's substantial
                                                        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_file:
                                                            tmp_file.write(image_data)
                                                            tmp_path = tmp_file.name
                                                        
                                                        print(f"Signature extraction: Saved image {obj_name} ({len(image_data)} bytes) to {tmp_path}")
                                                        
                                                        if image_count < 2:
                                                            key = 'signatory' if image_count == 0 else 'nc_representative'
                                                            signatures[key] = tmp_path
                                                            image_count += 1
                                                            print(f"Signature extraction: Assigned image as {key}")
                                                except Exception as e:
                                                    print(f"Error writing image {obj_name} to temp file: {e}")
                                            else:
                                                print(f"Signature extraction: Image {obj_name} has unsupported filter: {filter_type}")
                                    except Exception as e:
                                        print(f"Error processing image object {obj_name}: {e}")
                    except Exception as e:
                        print(f"Error processing page {page_num}: {e}")
                        continue
                
                print(f"Signature extraction: pypdf found {total_images_found} total images, extracted {len(signatures)} as signatures")
            except Exception as e:
                print(f"Error extracting images with pypdf: {e}")
                import traceback
                traceback.print_exc()
                
    except Exception as e:
        print(f"Error extracting signatures: {e}")
        import traceback
        traceback.print_exc()
        # Return empty dict on error - don't break the build
    
    if signatures:
        print(f"Signature extraction: Successfully extracted {len(signatures)} signatures: {list(signatures.keys())}")
    else:
        print("Signature extraction: No signatures were extracted")
    
    return {}  # Always return empty - signature extraction disabled

def create_service_agreement_from_data(csv_data, output_path, contact_name=None, source_pdf_path=None):
    """
    Create a service agreement PDF from provided data dictionary.
    
    Args:
        csv_data: Dictionary containing form data
        output_path: Path where the PDF should be saved
        contact_name: Optional name to use for Key Contact lookup
        source_pdf_path: Optional path to source PDF for signature extraction
    """
    # Load NDIS support items
    ndis_items = load_ndis_support_items()
    
    # Get team value to determine which active users CSV to use
    team_value = csv_data.get('Neighbourhood Care representative team', '')
    # Clean up checkbox characters
    team_value = team_value.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
    
    # Load active users based on team
    active_users = load_active_users(team_value)
    
    # Signature extraction removed to prevent timeouts
    signatures = {}
    
    # Create PDF document
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    _build_service_agreement_content(doc, csv_data, ndis_items, active_users, contact_name, signatures)

def create_service_agreement():
    # Load NDIS support items
    ndis_items = load_ndis_support_items()
    
    # Read data from PDF (preferred) or CSV (fallback)
    csv_data = {}
    pdf_path = 'outputs/other/Neighbourhood Care Welcoming Form Template 2.pdf'
    
    # Try to parse PDF first
    if os.path.exists(pdf_path):
        try:
            csv_data = parse_pdf_to_data(pdf_path)
            print(f"Successfully parsed PDF: {pdf_path}")
        except Exception as e:
            print(f"Error parsing PDF: {e}. Falling back to CSV.")
            csv_data = {}
    
    # Fallback to CSV if PDF parsing failed or didn't get enough data
    if not csv_data or len([v for v in csv_data.values() if v]) < 5:
        csv_candidates = [
            'outputs/other/Neighbourhood Care Welcoming Form Template 2.csv',
            'outputs/other/Neighbourhood Care Welcoming Form.csv'
        ]
        last_err = None
        for candidate in csv_candidates:
            try:
                with open(candidate, 'r', encoding='utf-8') as file:
                    reader = csv.DictReader(file)
                    for row in reader:
                        csv_data = row
                        break
                print(f"Successfully loaded CSV: {candidate}")
                break
            except Exception as e:
                last_err = e
                continue
        if not csv_data and last_err:
            raise last_err
    
    # Get team value to determine which active users CSV to use
    team_value = csv_data.get('Neighbourhood Care representative team', '')
    # Clean up checkbox characters
    team_value = team_value.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
    
    # Load active users based on team
    active_users = load_active_users(team_value)
    
    # Create PDF document
    doc = SimpleDocTemplate("Service Agreement - FINAL TABLES.pdf", pagesize=A4)
    _build_service_agreement_content(doc, csv_data, ndis_items, active_users)

def _add_header_footer(canvas_obj, doc):
    """Add header and footer to PDF pages"""
    # Footer color (gray)
    footer_color = colors.HexColor('#7F7F7F')
    
    # Footer text
    footer_text = "Neighbourhood Care | Suite 103, 19 Ogilvie Road, Mount Pleasant, WA 6153 | ABN 40 634 832 607"
    
    # Get page number
    page_num = canvas_obj.getPageNumber()
    
    # Footer position - center text at normal footer position
    footer_y = 30  # Position for footer text
    
    # Draw footer text in center
    canvas_obj.saveState()
    canvas_obj.setFillColor(footer_color)
    canvas_obj.setFont("Helvetica", 8)
    
    # Center footer text
    page_width = A4[0]
    footer_x_center = page_width / 2
    canvas_obj.drawCentredString(footer_x_center, footer_y, footer_text)
    
    # Page number on right side (a bit below the footer text)
    page_num_text = str(page_num)
    page_num_y = footer_y - 12  # Position below the footer text
    page_num_x = page_width - 50  # Right side with margin
    canvas_obj.drawRightString(page_num_x, page_num_y, page_num_text)
    
    canvas_obj.restoreState()

def _add_first_page_header(canvas_obj, doc):
    """Add header with image to first page only"""
    # Add the image to the right side of header
    image_filename = 'image.png'
    
    # Get the script directory (where this file is located)
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except:
        script_dir = os.getcwd()
    
    # Search for the image file
    image_path = None
    search_dirs = [script_dir, os.getcwd(), '.']
    
    for search_dir in search_dirs:
        if os.path.exists(search_dir):
            try:
                # Try exact filename first
                test_path = os.path.join(search_dir, image_filename)
                if os.path.exists(test_path):
                    image_path = os.path.abspath(test_path)
                    print(f"DEBUG: Found image file: {image_filename}")
                    break
                
                # Fallback: search for any file with "image" in the name
                for filename in os.listdir(search_dir):
                    if filename.lower() == image_filename.lower() or (filename.lower().startswith('image') and filename.lower().endswith('.png')):
                        full_path = os.path.join(search_dir, filename)
                        if os.path.exists(full_path):
                            image_path = os.path.abspath(full_path)
                            print(f"DEBUG: Found image file: {filename}")
                            break
                if image_path:
                    break
            except Exception as e:
                print(f"DEBUG: Error searching in {search_dir}: {e}")
                continue
    
    print(f"DEBUG: Looking for image: {image_filename}")
    print(f"DEBUG: Image found: {image_path}")
    print(f"DEBUG: Image exists: {os.path.exists(image_path) if image_path else False}")
    
    if os.path.exists(image_path):
        try:
            # Image size - doubled from original size
            # Body text is 11pt, so image height around 80 points for better visibility
            img_height = 80  # Doubled size for better visibility
            
            # Position on right side of header
            # ReportLab uses bottom-left as origin (0,0), so y increases upward
            page_width = A4[0]  # 595.27 points
            page_height = A4[1]  # 841.89 points
            img_y = page_height - 70  # Top of page with margin (from bottom)
            
            # Try to get image dimensions to calculate aspect ratio
            img_width = img_height * 1.5  # Default aspect ratio
            try:
                from PIL import Image as PILImage
                pil_img = PILImage.open(image_path)
                img_width_orig, img_height_orig = pil_img.size
                aspect_ratio = img_width_orig / img_height_orig
                img_width = img_height * aspect_ratio
                print(f"DEBUG: Image dimensions: {img_width_orig}x{img_height_orig}, scaled to: {img_width:.1f}x{img_height}")
            except Exception as pil_error:
                print(f"DEBUG: PIL not available, using default aspect ratio. Error: {pil_error}")
            
            img_x = page_width - img_width - 50  # Right side with margin
            print(f"DEBUG: Drawing image at position: x={img_x:.1f}, y={img_y}, size: {img_width:.1f}x{img_height}")
            print(f"DEBUG: Page dimensions: {page_width}x{page_height}")
            
            canvas_obj.saveState()
            # Use absolute path for drawImage
            abs_image_path = os.path.abspath(image_path)
            print(f"DEBUG: Using absolute image path: {abs_image_path}")
            
            # Draw the image
            canvas_obj.drawImage(
                abs_image_path, 
                img_x, 
                img_y, 
                width=img_width, 
                height=img_height, 
                preserveAspectRatio=True
            )
            canvas_obj.restoreState()
            print("DEBUG: Image drawn successfully")
        except Exception as e:
            print(f"ERROR: Could not add header image: {e}")
            import traceback
            traceback.print_exc()
    else:
        print(f"ERROR: Image file not found at: {image_path}")
        print(f"DEBUG: Current working directory: {os.getcwd()}")
        print(f"DEBUG: Script directory: {os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else 'unknown'}")
    
    # Also add footer for first page
    _add_header_footer(canvas_obj, doc)

def _build_service_agreement_content(doc, csv_data, ndis_items, active_users, contact_name=None, signatures=None):
    """Build the service agreement PDF content"""
    story = []
    styles = getSampleStyleSheet()
    
    # Create custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontSize=18,
        textColor=BLUE_COLOR,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=11,
        alignment=TA_LEFT,
        spaceAfter=0,
        leading=14,
        leftIndent=0
    )
    
    # Style with no space after for immediate following text
    normal_no_space_style = ParagraphStyle(
        'CustomNormalNoSpace',
        parent=styles['Normal'],
        fontSize=11,
        alignment=TA_LEFT,
        spaceAfter=0,
        leading=14,
        leftIndent=0
    )
    
    # Style for headings that should have no space after
    heading_no_space_style = ParagraphStyle(
        'CustomHeadingNoSpace',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=BLUE_COLOR,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0
    )
    
    # Style for black headings with no space after
    black_heading_no_space_style = ParagraphStyle(
        'BlackHeadingNoSpace',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.black,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0
    )
    
    # Style for bold headings (questions) with no space after
    bold_heading_no_space_style = ParagraphStyle(
        'BoldHeadingNoSpace',
        parent=styles['Normal'],
        fontSize=11,
        alignment=TA_LEFT,
        spaceAfter=0,
        leading=14,
        leftIndent=0
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=BLUE_COLOR,
        alignment=TA_LEFT,
        spaceAfter=12,
        leftIndent=0
    )
    
    black_heading_style = ParagraphStyle(
        'BlackHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.black,
        alignment=TA_LEFT,
        spaceAfter=12,
        leftIndent=0
    )
    
    table_text_style = ParagraphStyle(
        'TableText',
        wordWrap='CJK',  # Enable word wrapping
        parent=styles['Normal'],
        fontSize=8,
        alignment=TA_LEFT,
        spaceAfter=0,
        leading=10,
        leftIndent=0
    )
    
    bullet_style = ParagraphStyle(
        'BulletStyle',
        parent=styles['Normal'],
        fontSize=11,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=20,
        bulletIndent=10,
        leading=14
    )
    
    # Get team value early for price determination
    team_value = csv_data.get('Neighbourhood Care representative team', '[To be filled in]')
    # Clean up checkbox characters that appear as black boxes
    team_value = team_value.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
    
    # Title
    story.append(Paragraph("Service Agreement", title_style))
    story.append(Spacer(1, 12))  # Add one line space after title
    
    # Introduction
    intro1 = "Thank you for choosing Neighbourhood Care. We look forward to working with you to help you achieve your goals."
    story.append(Paragraph(intro1, normal_no_space_style))
    story.append(Spacer(1, 12))
    
    intro2 = "This document is a written agreement between you and Neighbourhood Care that outlines the supports we will provide and how they will be delivered."
    story.append(Paragraph(intro2, normal_no_space_style))
    story.append(Spacer(1, 12))
    
    intro3 = "<b>Please make sure you have read and understood our Agreements, Promises and Terms of Service before completing this document.</b>"
    story.append(Paragraph(intro3, normal_no_space_style))
    story.append(Spacer(1, 12))
    
    intro4 = "If you are unsure about any part of this document please speak to your Neighbourhood Care representative."
    story.append(Paragraph(intro4, normal_no_space_style))
    story.append(Spacer(1, 12))
    
    intro5 = "This Service Agreement must then be signed for us to start delivering services."
    story.append(Paragraph(intro5, normal_no_space_style))
    story.append(Spacer(1, 12))
    
    # What makes up your service
    what_makes_up_heading_style = ParagraphStyle(
        'WhatMakesUpHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=BLUE_COLOR,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0
    )
    story.append(Paragraph("What makes up your service?", what_makes_up_heading_style))
    
    service_text = "Please note that your service is made up of face to face and some non face to face supports. Services that may be charged as part of your service are:"
    story.append(Paragraph(service_text, normal_no_space_style))
    story.append(Spacer(1, 12))
    
    service_bullets = [
        "Transporting you during a shift (this is a $1 cost per km and is billed out of your core budget).",
        "Communication by phone or email or in a face to face meeting with key people in your network - when this is not part of your rostered shift.",
        "Travel for support workers or therapists when they are coming directly from the office or from another participant or travelling back to the office at the end of the shift.",
        "Preparing some reports that are required for the NDIS such as creating your Support Plan.",
        "Costs for when we are supporting you in the community such as parking, public transport and so forth.",
        "For <i>new</i> participants, receiving Core supports, the one off Establishment fee is applied."
    ]
    
    for bullet in service_bullets:
        story.append(Paragraph(f"• {bullet}", bullet_style))
    
    story.append(Spacer(1, 12))
    
    # Calculate Establishment Fee
    establishment_fee_amount = get_establishment_fee(csv_data, ndis_items, team_value)
    
    # Only show establishment fee table if fee is greater than $0.00
    if establishment_fee_amount and establishment_fee_amount != '$0.00':
        # Extract numeric value for comparison
        try:
            fee_value = float(establishment_fee_amount.replace('$', '').replace(',', ''))
            if fee_value > 0:
                # Establishment Fee
                establishment_data = [
                    ['Establishment Fee', establishment_fee_amount]
                ]
                
                establishment_table = Table(establishment_data, colWidths=[3*inch, 2*inch])
                establishment_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (0, 0), BLUE_COLOR),
                    ('TEXTCOLOR', (0, 0), (0, 0), colors.white),
                    ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
                    ('BACKGROUND', (1, 0), (1, 0), colors.white),
                    ('TEXTCOLOR', (1, 0), (1, 0), colors.black),
                    ('FONTNAME', (1, 0), (1, 0), 'Helvetica'),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('TOPPADDING', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
                ]))
                story.append(establishment_table)
                story.append(Spacer(1, 12))
        except (ValueError, AttributeError):
            # If we can't parse the fee, don't show the table
            pass
    
    # Schedule of Supports
    schedule_heading_style = ParagraphStyle(
        'ScheduleHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=BLUE_COLOR,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0
    )
    story.append(Paragraph("Schedule of Supports", schedule_heading_style))
    
    # Core and Capacity Building
    core_capacity_heading_style = ParagraphStyle(
        'CoreCapacityHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.black,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0
    )
    story.append(Paragraph("Core and Capacity Building", core_capacity_heading_style))
    # Create white bold text style for table cells
    white_bold_table_text_style = ParagraphStyle(
        'WhiteBoldTableText',
        parent=styles['Normal'],
        fontSize=8,
        alignment=TA_LEFT,
        spaceAfter=0,
        leading=10,
        leftIndent=0,
        textColor=colors.white,
        fontName='Helvetica-Bold'
    )
    
    core_data = [
        [Paragraph('Core Budget Allocated to Neighbourhood Care', white_bold_table_text_style), Paragraph(csv_data.get('Total core budget to allocate to Neighbourhood Care', 'Total core budget to allocate to Neighbourhood Care (NDIS Information)'), table_text_style)],
        [Paragraph('Capacity Building Budget Allocated to Neighbourhood Care', white_bold_table_text_style), Paragraph(csv_data.get('Total capacity building budget to allocate to Neighbourhood Care', 'Total capacity building budget to allocate to Neighbourhood Care (NDIS Information)'), table_text_style)]
    ]
    
    core_table = Table(core_data, colWidths=[3.5*inch, 2*inch])
    core_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), BLUE_COLOR),
        ('TEXTCOLOR', (0, 0), (0, 0), colors.white),
        ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
        ('BACKGROUND', (1, 0), (1, 0), colors.white),
        ('TEXTCOLOR', (1, 0), (1, 0), colors.black),
        ('FONTNAME', (1, 0), (1, 0), 'Helvetica'),
        ('BACKGROUND', (0, 1), (0, 1), BLUE_COLOR),
        ('TEXTCOLOR', (0, 1), (0, 1), colors.white),
        ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Bold'),
        ('BACKGROUND', (1, 1), (1, 1), colors.white),
        ('TEXTCOLOR', (1, 1), (1, 1), colors.black),
        ('FONTNAME', (1, 1), (1, 1), 'Helvetica'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    story.append(core_table)
    
    # Support Items
    support_items_heading_style = ParagraphStyle(
        'SupportItemsHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.black,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0
    )
    story.append(Paragraph("Support Items", support_items_heading_style))
    support_data = [['Category', 'Name', 'Number', 'Unit', 'Price']]
    
    # Extract support items from the PDF data - look for "Support item (X) (Support Items Required)"
    support_items_from_pdf = []
    for i in range(1, 20):  # Check up to 20 support items
        key = f'Support item ({i}) (Support Items Required)'
        item_name = csv_data.get(key, '').strip()
        if item_name:
            support_items_from_pdf.append((i, item_name))
    
    # If no support items found in PDF, use empty list (don't show hardcoded items)
    # Determine which state's price to use based on team
    price_state = get_price_state(team_value)
    price_key = 'wa_price' if price_state == 'WA' else 'qld_price'
    
    for item_num, item_name in support_items_from_pdf:
        item_details = lookup_support_item(ndis_items, item_name)
        # Check if item was actually found (not the placeholder)
        item_found = item_name in ndis_items or any(
            item_name.lower() in key.lower() or key.lower() in item_name.lower() 
            for key in ndis_items.keys()
        )
        # If item not found, show [Not Found] for all fields
        if item_found:
            support_data.append([
                Paragraph(f'Support item ({item_num})', table_text_style),
                Paragraph(item_name, table_text_style),
                Paragraph(item_details.get('number', ''), table_text_style),
                Paragraph(item_details.get('unit', ''), table_text_style),
                Paragraph(item_details.get(price_key, ''), table_text_style)
            ])
        else:
            support_data.append([
                Paragraph(f'Support item ({item_num})', table_text_style),
                Paragraph(item_name, table_text_style),
                Paragraph('[Not Found]', table_text_style),
                Paragraph('[Not Found]', table_text_style),
                Paragraph('[Not Found]', table_text_style)
            ])
    
    # Adjust column widths to prevent text overflow - A4 width is ~8.27 inches, leave some margin
    support_table = Table(support_data, colWidths=[0.7*inch, 3.5*inch, 1.1*inch, 0.7*inch, 0.9*inch])
    support_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLUE_COLOR),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    story.append(support_table)
    story.append(Spacer(1, 12))
    
    # Consents
    consents_heading_style = ParagraphStyle(
        'ConsentsHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=BLUE_COLOR,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0
    )
    story.append(Paragraph("Consents", consents_heading_style))
    consent_data = []
    
    consents = [
        'I agree to receive services from Neighbourhood Care.',
        'I consent for Neighbourhood Care to create an NDIS portal service booking on my behalf if my budget/s are Agency Managed.',
        'I understand that if at any time I (The Participant) require emergency medical assistance, Neighbourhood Care staff will call an ambulance to attend, and that I (The Participant) will be liable for any expenses incurred for Ambulance attendance.',
        'I agree that Neighbourhood Care staff may administer simple first aid to me (The Participant), if the need arises.',
        'I consent for Neighbourhood Care to discuss relevant information about my case with other providers involved in my care and support, for example GP, support coordinator.',
        'I agree not to smoke inside the home whilst Neighbourhood Care staff are present.',
        'I understand that an Emergency Response Plan will be developed with me by Neighbourhood Care to help keep me safe in the event of an emergency.',
        'I consent for Neighbourhood Care for I (The Participant) to be photographed/recorded for therapeutic and/or training purposes.',
        'I give authority for my details or information to be shared with an external auditor who will assess Neighbourhood Care against the NDIS Quality and Safeguards Framework.'
    ]
    
    for consent in consents:
        consent_data.append([Paragraph(consent, white_bold_table_text_style), csv_data.get(consent, 'Yes')])
    
    consent_table = Table(consent_data, colWidths=[4.2*inch, 0.8*inch])
    consent_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), BLUE_COLOR),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.white),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('BACKGROUND', (1, 0), (1, -1), colors.white),
        ('TEXTCOLOR', (1, 0), (1, -1), colors.black),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    story.append(consent_table)
    story.append(Spacer(1, 12))
    
    # Agreements, Promises and Terms of Service
    agreements_heading_style = ParagraphStyle(
        'AgreementsHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=BLUE_COLOR,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0
    )
    story.append(Paragraph("<b>Agreements, Promises and Terms of Service</b>", agreements_heading_style))
    
    story.append(Paragraph("Our Agreements, Promises and Terms of Service outline how we deliver services. It outlines our rights and responsibilities as a service provider, and the rights and responsibilities of the people we provide services to.", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>What can you expect from Neighbourhood Care?</b>", bold_heading_no_space_style))
    story.append(Paragraph("We agree to:", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    nc_agreements = [
        "Review your care and service plan every 6 months with you.",
        "Maintain a service that works for you, so times of appointments meet your needs and we are in tune with each other. We call this Attunement.",
        "At all times communicate openly and honestly in a timely manner.",
        "At all times treat you with dignity and respect and being mindful of any cultural differences.",
        "Be open and transparent about managing complaints or disagreements and provide you the opportunity to provide feedback to us and to the NDIS.",
        "Ensure your privacy and any information is held in confidence and not shared without your permission.",
        "Work together at every step on your journey towards reaching your goals.",
        "Operate within the National Disability Insurance Scheme Act 2013 and associated Business Rules."
    ]
    
    for agreement in nc_agreements:
        story.append(Paragraph(f"• {agreement}", bullet_style))
    
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>What is expected of you as an NDIS participant?</b>", normal_style))
    story.append(Paragraph("You agree to:", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    participant_agreements = [
        "Inform Neighbourhood Care about how you wish your supports to be provided and how they should be offered to meet your needs.",
        "Treat Neighbourhood Care staff with courtesy and respect in the same way you want to be treated.",
        "Talk to Neighbourhood Care if you have any concerns about Plan Management or Financial Administration being provided.",
        "Give your care and support team the required notice if you need to end this Service Agreement. There is a notice period of 4 weeks to end this service.",
        "Advise your care and support team immediately if your plan is suspended or replaced by a new NDIS Plan or where you stop being an active participant in the NDIS."
    ]
    
    for agreement in participant_agreements:
        story.append(Paragraph(f"• {agreement}", bullet_style))
    
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>Cancellations</b>", normal_style))
    story.append(Paragraph("If a make-up shift with that support worker cannot be scheduled, the NDIS considers this a Short Notice Cancellation, and Neighbourhood Care may charge 100% of the agreed hourly rate.", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>How will services be provided to you?</b>", bold_heading_no_space_style))
    story.append(Paragraph("Services will be provided at your place of residence and in other locations as deemed necessary and suitable by you, your family, the Support Coordinator and the Neighbourhood Care Team charged with your safety whilst in the service.", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>When will services be provided?</b>", bold_heading_no_space_style))
    story.append(Paragraph("All services will be provided in attunement to your needs and subject to availability by others who may have an impact to your availability.", normal_no_space_style))
    story.append(Paragraph("From the commencement of the service agreement: Direct support provided as per the support and/or Therapy support plan, subject to change/increase, upon confirmation with you and/or your family.", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>How long will services be provided?</b>", bold_heading_no_space_style))
    story.append(Paragraph("Services will be provided for the length of the service agreement plan unless otherwise ceased at the discretion by by you or your team, in accordance with Neighbourhood Care's Policy and Procedures.", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>How to make changes?</b>", bold_heading_no_space_style))
    story.append(Paragraph("If changes to the supports or their delivery are required, you and your Neighbourhood Care team (the parties) agree to discuss and review this Service Agreement. The Parties agree that any changes to this Service Agreement will be in writing, signed, and dated by the Parties.", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>How to end the Agreement?</b>", bold_heading_no_space_style))
    story.append(Paragraph("Should either Party wish to end this Service Agreement they must give 4 weeks written notice to their care and support team. If either Party seriously breaches this Service Agreement the requirement of notice will be waived.", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>Pricing Changes</b>", bold_heading_no_space_style))
    story.append(Paragraph("Neighbourhood Care's services are charged in accordance with the NDIS Pricing Arrangements and Price Limits Guide. The prices set out in this Service Agreement will change in accordance with updates to the NDIS Pricing Arrangements and Price Limits Guide. This typically updates on the 1st of July each year but may be updated at other times.", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>What to do if there is a problem?</b>", bold_heading_no_space_style))
    story.append(Paragraph("If there is a problem with anything related to your service or this agreement, you can contact:", normal_no_space_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Your Neighbourhood Care contact person (please refer to the front page of your Service Agreement) or 1800 292 273.", normal_no_space_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Alternatively, you can email your concern or query to: ask@nhcare.com.au.", normal_no_space_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph("If you don't feel that your problem was resolved please speak to your support coordinator, Local Area Coordinator or you can just contact the National Disability Insurance Agency (NDIA)", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>Collection of your personal information</b>", bold_heading_no_space_style))
    story.append(Paragraph("Neighbourhood Care will use your information to support your involvement in the NDIS.", normal_no_space_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Neighbourhood Care will NOT use any of your personal information for any other purpose or disclose your personal information to any other organisations or individuals (including overseas recipients) unless authorised by law or you provide consent for us to do so.", normal_no_space_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph("You can also ask to see what personal information (if any) we hold about you at any time and you can seek correction if the information is incorrect.", normal_no_space_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph("<b>Neighbourhood Care's privacy policy describes:</b>", normal_no_space_style))
    
    privacy_bullets = [
        "How we use your personal information",
        "Why some personal information may be given to other organisations from time to time",
        "How you can access the personal information we have about you on our system",
        "How you can complain about a privacy breach, and how Neighbourhood Care deals with the complaint.",
        "How you can get your personal information corrected if it is wrong."
    ]
    
    for bullet in privacy_bullets:
        story.append(Paragraph(f"• {bullet}", bullet_style))
    
    story.append(Spacer(1, 12))
    story.append(Paragraph("You can find the policy by enquiring at Neighbourhood Care.", normal_no_space_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Please note that Neighbourhood Care is required to release information about service users (without identifying you by full name or address) to the Australian Institute of Health and Welfare, to enable statistics about disability services and their clients to be compiled. This information will be kept confidential. This information is used for statistical purposes only and will not be used to affect your entitlements or your access to services. You have the right to access your own files and to update or correct information included in the Disability Services National Minimum Data Set collection.", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    story.append(Paragraph("<b>Goods and Services Tax</b>", normal_style))
    story.append(Paragraph("Most services provided under the NDIS will not include GST. However, GST will apply to some services. Neighbourhood Care will apply GST when it is required.", normal_no_space_style))
    story.append(Spacer(1, 12))
    
    # For your information text (comes before Signatures heading)
    for_your_info_text = 'For your information: "A supply of supports under this Service Agreement is a supply of one or more reasonable and necessary supports specified in the statement of supports included, under subsection 33(2) of the National Disability Insurance Scheme Act 2013 (NDIS Act), in the participant\'s NDIS Plan currently in effect under section 37 of the NDIS Act."'
    story.append(Paragraph(for_your_info_text, normal_no_space_style))
    story.append(Spacer(1, 12))
    
    # Signatures
    signatures_heading_style = ParagraphStyle(
        'SignaturesHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=BLUE_COLOR,
        alignment=TA_LEFT,
        spaceAfter=0,
        leftIndent=0
    )
    story.append(Paragraph("Signatures", signatures_heading_style))
    signatory_name = f"{csv_data.get('First name (Person Signing the Agreement)', 'First name (Person Signing the Agreement)')} {csv_data.get('Surname (Person Signing the Agreement)', 'Surname (Person Signing the Agreement)')}"
    signatory_text = f"<b>Signatory:</b><br/><b>Name:</b> {signatory_name}<br/><b>Date:</b> [Date]<br/><b>Signed:</b>"
    story.append(Paragraph(signatory_text, normal_no_space_style))
    
    # Add signatory signature image if available
    if signatures and 'signatory' in signatures:
        try:
            sig_img = signatures['signatory']
            # Check if it's a file path (string) or already an image object
            if isinstance(sig_img, str) and os.path.exists(sig_img):
                print(f"Adding signatory signature from: {sig_img}")
                story.append(Spacer(1, 6))
                story.append(Image(sig_img, width=2*inch, height=0.5*inch))
            else:
                print(f"Signatory signature path invalid or doesn't exist: {sig_img}")
                story.append(Paragraph("[Signature]", normal_no_space_style))
        except Exception as e:
            print(f"Error adding signatory signature: {e}")
            import traceback
            traceback.print_exc()
            story.append(Paragraph("[Signature]", normal_no_space_style))
    else:
        print(f"Signatory signature not found. Available signatures: {list(signatures.keys()) if signatures else 'none'}")
        story.append(Paragraph("[Signature]", normal_no_space_style))
    
    story.append(Spacer(1, 12))
    
    # Neighbourhood Care Representative
    nc_rep_text = f"<b>Neighbourhood Care Representative:</b><br/><b>Name:</b> [To be filled with NC representative name]<br/><b>Date:</b> [Date]<br/><b>Signed:</b>"
    story.append(Paragraph(nc_rep_text, normal_no_space_style))
    
    # Add NC representative signature image if available
    if signatures and 'nc_representative' in signatures:
        try:
            sig_img = signatures['nc_representative']
            # Check if it's a file path (string) or already an image object
            if isinstance(sig_img, str) and os.path.exists(sig_img):
                print(f"Adding NC representative signature from: {sig_img}")
                story.append(Spacer(1, 6))
                story.append(Image(sig_img, width=2*inch, height=0.5*inch))
            else:
                print(f"NC representative signature path invalid or doesn't exist: {sig_img}")
                story.append(Paragraph("[Signature]", normal_no_space_style))
        except Exception as e:
            print(f"Error adding NC representative signature: {e}")
            import traceback
            traceback.print_exc()
            story.append(Paragraph("[Signature]", normal_no_space_style))
    else:
        print(f"NC representative signature not found. Available signatures: {list(signatures.keys()) if signatures else 'none'}")
        story.append(Paragraph("[Signature]", normal_no_space_style))
    
    story.append(Spacer(1, 12))
    
    # Participant - FIXED with all missing fields
    story.append(Paragraph("Appendix", black_heading_style))
    story.append(Paragraph("Participant", black_heading_no_space_style))
    # Participant Name: First name + Middle name + Surname (from Details of the Client)
    first_name = csv_data.get('First name (Details of the Client)', '').strip()
    middle_name = csv_data.get('Middle name (Details of the Client)', '').strip()
    surname = csv_data.get('Surname (Details of the Client)', '').strip()
    participant_name_parts = [p for p in [first_name, middle_name, surname] if p]
    participant_name = ' '.join(participant_name_parts) if participant_name_parts else ''
    
    # Emergency Contact: First name + Surname (from Emergency contact)
    emergency_first = csv_data.get('First name (Emergency contact)', '').strip()
    emergency_surname = csv_data.get('Surname (Emergency contact)', '').strip()
    emergency_contact_parts = [p for p in [emergency_first, emergency_surname] if p]
    emergency_contact = ' '.join(emergency_contact_parts) if emergency_contact_parts else get_emergency_contact(csv_data)
    
    # Get all values, using empty string if not found
    home_address = csv_data.get('Home address (Contact Details of the Client)', '').strip()
    home_phone = csv_data.get('Home phone (Contact Details of the Client)', '').strip()
    mobile_phone = csv_data.get('Mobile phone (Contact Details of the Client)', '').strip()
    email_address = csv_data.get('Email address (Contact Details of the Client)', '').strip()
    dob = csv_data.get('Date of birth (Details of the Client)', '').strip() or csv_data.get('Date of birth', '').strip()
    ndis_num = csv_data.get('NDIS number (Details of the Client)', '').strip() or csv_data.get('NDIS number', '').strip()
    plan_start = csv_data.get('Plan start date', '').strip()
    plan_end = csv_data.get('Plan end date', '').strip()
    service_start = csv_data.get('Service start date', '').strip() or csv_data.get('Service start', '').strip()
    service_end = csv_data.get('Service end date', '').strip() or csv_data.get('Service end', '').strip()
    preferred_contact = csv_data.get('Preferred method of contact', '').strip()
    # Clean up checkbox characters that appear as black boxes
    preferred_contact = preferred_contact.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
    
    participant_data = [
        ['Participant Name', Paragraph(participant_name, table_text_style)],
        ['Date of Birth', dob],
        ['NDIS Number', ndis_num],
        ['Plan Duration', f"{plan_start} - {plan_end}" if (plan_start or plan_end) else ''],
        ['Address', Paragraph(home_address, table_text_style)],
        ['Home phone', home_phone],
        ['Mobile phone', mobile_phone],
        ['Email address', Paragraph(email_address, table_text_style)],
        ['Preferred contact method', preferred_contact],
        ['Emergency Contact', Paragraph(emergency_contact, table_text_style)],
        ['Service Agreement Duration', f"{service_start} - {service_end}" if (service_start or service_end) else '']
    ]
    
    participant_table = Table(participant_data, colWidths=[2.5*inch, 3*inch])
    participant_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    story.append(participant_table)
    story.append(Spacer(1, 12))
    
    # Signatory (detailed) - FIXED with all missing fields
    story.append(Paragraph("Signatory", black_heading_no_space_style))
    # Get signatory contact details based on who is signing (preferred method only)
    signatory_contact = get_signatory_contact_details(csv_data)
    
    signatory_detailed_data = [
        ['Name', Paragraph(get_signatory_name(csv_data), table_text_style)],
        ['Relationship to Participant', get_signatory_relationship(csv_data)],
        ['Address', Paragraph(get_signatory_address(csv_data), table_text_style)],
        ['Contact Details', Paragraph(signatory_contact, table_text_style)]
    ]
    
    signatory_detailed_table = Table(signatory_detailed_data, colWidths=[2.5*inch, 3*inch])
    signatory_detailed_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    story.append(signatory_detailed_table)
    story.append(Spacer(1, 12))
    
    # Plan Manager
    story.append(Paragraph("Plan Manager", black_heading_no_space_style))
    plan_manager_data = [
        ['Name', get_plan_manager_name(csv_data)],
        ['Postal Address', Paragraph(get_plan_manager_address(csv_data), table_text_style)],
        ['Phone', get_plan_manager_phone(csv_data)],
        ['Email Address', Paragraph(get_plan_manager_email(csv_data), table_text_style)]
    ]
    
    plan_manager_table = Table(plan_manager_data, colWidths=[2.5*inch, 3*inch])
    plan_manager_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    story.append(plan_manager_table)
    story.append(Spacer(1, 12))
    
    # My Neighbourhood Care Key Contact
    story.append(Paragraph("My Neighbourhood Care Key Contact", black_heading_no_space_style))
    
    # Get contact name from parameter or fallback to Respondent field
    contact_name_to_use = contact_name or csv_data.get('Respondent', '')
    user_data = lookup_user_data(active_users, contact_name_to_use) if contact_name_to_use else {'name': '', 'mobile': '', 'email': ''}
    
    # Calculate My Neighbourhood Care ID: First name + Surname + Year of Date of birth
    first_name = csv_data.get('First name (Details of the Client)', '').strip()
    surname = csv_data.get('Surname (Details of the Client)', '').strip()
    dob = csv_data.get('Date of birth (Details of the Client)', '').strip()
    
    # Extract year from date of birth (handle formats like DD/MM/YYYY or YYYY-MM-DD)
    year = ''
    if dob:
        # Try to extract year from common date formats
        year_match = re.search(r'\b(19|20)\d{2}\b', dob)
        if year_match:
            year = year_match.group(0)
    
    # Build ID: First name + Surname + Year (with spaces)
    name_parts = [p for p in [first_name, surname] if p]
    neighbourhood_care_id = ' '.join(name_parts) + ' ' + year if name_parts and year else '[To be filled in]'
    
    # Get team value and clean checkbox characters
    team_value = csv_data.get('Neighbourhood Care representative team', '[To be filled in]')
    # Clean up checkbox characters that appear as black boxes
    team_value = team_value.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
    
    key_contact_data = [
        ['My Neighbourhood Care ID', neighbourhood_care_id],
        ['Team', team_value],
        ['Key Contact', Paragraph(contact_name_to_use if contact_name_to_use else user_data.get('name', '[To be filled in]'), table_text_style)],
        ['Phone', user_data.get('mobile', '[To be filled in]')],
        ['Email Address', Paragraph(user_data.get('email', '[To be filled in]'), table_text_style)],
        ['Neighbourhood Care Office', 'Phone: 1800 292 273']
    ]
    
    key_contact_table = Table(key_contact_data, colWidths=[2.5*inch, 3*inch])
    key_contact_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    story.append(key_contact_table)
    
    # Build PDF with headers and footers
    doc.build(story, onFirstPage=_add_first_page_header, onLaterPages=_add_header_footer)
    print("Service Agreement PDF FINAL TABLES created successfully!")

def get_emergency_contact(csv_data):
    """Get emergency contact based on the logic specified"""
    is_primary_carer = csv_data.get('Is the primary carer also the emergency contact for the participant?', '').strip()
    # Clean checkbox characters and check if it's "yes"
    is_primary_carer_clean = is_primary_carer.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip().lower()
    
    if 'yes' in is_primary_carer_clean:
        first_name = csv_data.get('First name (Primary carer)', '').strip()
        surname = csv_data.get('Surname (Primary carer)', '').strip()
        name_parts = [p for p in [first_name, surname] if p]
        return ' '.join(name_parts) if name_parts else ''
    else:
        first_name = csv_data.get('First name (Emergency contact)', '').strip()
        surname = csv_data.get('Surname (Emergency contact)', '').strip()
        name_parts = [p for p in [first_name, surname] if p]
        return ' '.join(name_parts) if name_parts else ''

def get_signatory_name(csv_data):
    """Get signatory name based on who is signing"""
    person_signing = csv_data.get('Person signing the agreement', '').strip()
    if person_signing.lower() == 'participant':
        # Participant is the client - use First name + Middle name + Surname from Details of the Client
        first_name = csv_data.get('First name (Details of the Client)', '').strip()
        middle_name = csv_data.get('Middle name (Details of the Client)', '').strip()
        surname = csv_data.get('Surname (Details of the Client)', '').strip()
        name_parts = [p for p in [first_name, middle_name, surname] if p]
        return ' '.join(name_parts) if name_parts else ''
    elif person_signing.lower() == 'primary carer':
        first_name = csv_data.get('First name (Primary carer)', '').strip()
        surname = csv_data.get('Surname (Primary carer)', '').strip()
        name_parts = [p for p in [first_name, surname] if p]
        return ' '.join(name_parts) if name_parts else ''
    else:
        first_name = csv_data.get('First name (Person Signing the Agreement)', '').strip()
        surname = csv_data.get('Surname (Person Signing the Agreement)', '').strip()
        name_parts = [p for p in [first_name, surname] if p]
        return ' '.join(name_parts) if name_parts else ''

def get_signatory_relationship(csv_data):
    """Get signatory relationship based on who is signing"""
    person_signing = csv_data.get('Person signing the agreement', '').strip()
    if person_signing.lower() == 'participant':
        return 'Participant'
    elif person_signing.lower() == 'primary carer':
        return csv_data.get('Relationship to client (Primary carer)', '').strip()
    else:
        return csv_data.get('Relationship to client (Person Signing the Agreement)', '').strip()

def get_signatory_address(csv_data):
    """Get signatory address based on who is signing"""
    person_signing = csv_data.get('Person signing the agreement', '').strip()
    if person_signing.lower() == 'participant':
        return csv_data.get('Home address (Contact Details of the Client)', '').strip()
    elif person_signing.lower() == 'primary carer':
        return csv_data.get('Home address (Primary carer)', '').strip()
    else:
        return csv_data.get('Home address (Person Signing the Agreement)', '').strip()

def get_signatory_contact_details(csv_data):
    """Get actual contact detail value for signatory based on preferred method and who is signing"""
    person_signing = csv_data.get('Person signing the agreement', '').strip()
    
    # Get preferred method of contact
    preferred_contact = ''
    if person_signing.lower() == 'participant':
        # Use participant's preferred method of contact
        preferred_contact = csv_data.get('Preferred method of contact', '').strip()
    elif person_signing.lower() == 'primary carer':
        # Use primary carer's preferred method of contact (if available)
        preferred_contact = csv_data.get('Preferred method of contact (Primary carer)', '').strip() or csv_data.get('Preferred method of contact', '').strip()
    else:
        # Use Person Signing the Agreement preferred method of contact (if available)
        preferred_contact = csv_data.get('Preferred method of contact (Person Signing the Agreement)', '').strip() or csv_data.get('Preferred method of contact', '').strip()
    
    # Clean up checkbox characters to get the actual preferred method
    if preferred_contact:
        preferred_contact = preferred_contact.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
    
    # Get the actual contact value based on preferred method
    preferred_contact_lower = preferred_contact.lower()
    
    if person_signing.lower() == 'participant':
        # Get participant's contact details
        if 'home phone' in preferred_contact_lower:
            return csv_data.get('Home phone (Contact Details of the Client)', '').strip()
        elif 'mobile phone' in preferred_contact_lower or 'mobile' in preferred_contact_lower:
            return csv_data.get('Mobile phone (Contact Details of the Client)', '').strip()
        elif 'email' in preferred_contact_lower:
            return csv_data.get('Email address (Contact Details of the Client)', '').strip()
        elif 'work phone' in preferred_contact_lower:
            return csv_data.get('Work phone (Contact Details of the Client)', '').strip()
    elif person_signing.lower() == 'primary carer':
        # Get primary carer's contact details
        if 'home phone' in preferred_contact_lower:
            return csv_data.get('Home phone (Primary carer)', '').strip()
        elif 'mobile phone' in preferred_contact_lower or 'mobile' in preferred_contact_lower:
            return csv_data.get('Mobile phone (Primary carer)', '').strip()
        elif 'email' in preferred_contact_lower:
            return csv_data.get('Email address (Primary carer)', '').strip()
        elif 'work phone' in preferred_contact_lower:
            return csv_data.get('Work phone (Primary carer)', '').strip()
    else:
        # Get Person Signing the Agreement's contact details
        if 'home phone' in preferred_contact_lower:
            return csv_data.get('Home phone (Person Signing the Agreement)', '').strip()
        elif 'mobile phone' in preferred_contact_lower or 'mobile' in preferred_contact_lower:
            return csv_data.get('Mobile phone (Person Signing the Agreement)', '').strip()
        elif 'email' in preferred_contact_lower:
            return csv_data.get('Email address (Person Signing the Agreement)', '').strip()
        elif 'work phone' in preferred_contact_lower:
            return csv_data.get('Work phone (Person Signing the Agreement)', '').strip()
    
    # Fallback: return preferred method if we can't find the actual value
    return preferred_contact

def get_plan_manager_name(csv_data):
    """Get plan manager name based on plan management type"""
    plan_type = csv_data.get('Plan management type', '')
    if plan_type in ['NDIA Agency Managed', 'Insurance Commission of WA']:
        return ''
    else:
        return csv_data.get('Plan manager name', 'Plan manager name (Support Items Required)')

def get_plan_manager_address(csv_data):
    """Get plan manager address based on plan management type"""
    plan_type = csv_data.get('Plan management type', '')
    if plan_type in ['NDIA Agency Managed', 'Insurance Commission of WA']:
        return ''
    else:
        return csv_data.get('Plan manager postal address', 'Plan manager postal address (Support Items Required)')

def get_plan_manager_phone(csv_data):
    """Get plan manager phone based on plan management type"""
    plan_type = csv_data.get('Plan management type', '')
    if plan_type in ['NDIA Agency Managed', 'Insurance Commission of WA']:
        return ''
    else:
        return csv_data.get('Plan manager phone number', 'Plan manager phone number (Support Items Required)')

def get_plan_manager_email(csv_data):
    """Get plan manager email based on plan management type"""
    plan_type = csv_data.get('Plan management type', '')
    if plan_type in ['NDIA Agency Managed', 'Insurance Commission of WA']:
        return ''
    else:
        return csv_data.get('Plan manager email address', 'Plan manager email address (Support Items Required)')

def format_date_for_display(date_str):
    """Format date string to DD/MM/YYYY format"""
    if not date_str:
        return ""
    
    date_str = date_str.strip()
    if not date_str:
        return ""
    
    from datetime import datetime
    
    # Try to parse common date formats
    date_formats = [
        '%Y-%m-%d',      # 2023-12-25
        '%d/%m/%Y',      # 25/12/2023
        '%m/%d/%Y',      # 12/25/2023
        '%d-%m-%Y',      # 25-12-2023
        '%Y/%m/%d',      # 2023/12/25
        '%d.%m.%Y',      # 25.12.2023
    ]
    
    for fmt in date_formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime('%d/%m/%Y')
        except ValueError:
            continue
    
    # If no format matched, try to extract numbers
    try:
        numbers = re.findall(r'\d+', date_str)
        if len(numbers) >= 3:
            if len(numbers[0]) == 4:
                year, month, day = numbers[0], numbers[1], numbers[2]
            elif len(numbers[-1]) == 4:
                day, month, year = numbers[0], numbers[1], numbers[2]
            else:
                day, month, year = numbers[0], numbers[1], numbers[2]
            
            day = day.zfill(2)
            month = month.zfill(2)
            return f"{day}/{month}/{year}"
    except (ValueError, IndexError, AttributeError):
        pass
    
    return date_str

def get_emergency_contact_phone(csv_data):
    """Get emergency contact phone numbers (Home phone + Mobile phone + Work phone)"""
    phones = []
    
    # Try to get emergency contact phone fields
    home_phone = csv_data.get('Home phone (Emergency contact)', '').strip()
    mobile_phone = csv_data.get('Mobile phone (Emergency contact)', '').strip()
    work_phone = csv_data.get('Work phone (Emergency contact)', '').strip()
    
    # If emergency contact fields don't exist, check if primary carer is emergency contact
    if not home_phone and not mobile_phone and not work_phone:
        is_primary_carer = csv_data.get('Is the primary carer also the emergency contact for the participant?', '').strip()
        is_primary_carer_clean = is_primary_carer.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip().lower()
        
        if 'yes' in is_primary_carer_clean:
            home_phone = csv_data.get('Home phone (Primary carer)', '').strip()
            mobile_phone = csv_data.get('Mobile phone (Primary carer)', '').strip()
            work_phone = csv_data.get('Work phone (Primary carer)', '').strip()
    
    if home_phone:
        phones.append(home_phone)
    if mobile_phone:
        phones.append(mobile_phone)
    if work_phone:
        phones.append(work_phone)
    
    return '; '.join(phones) if phones else ''

def get_emergency_contact_relationship(csv_data):
    """Get emergency contact relationship to client"""
    # First try the emergency contact relationship field
    relationship = csv_data.get('Relationship to client (Emergency contact)', '').strip()
    
    if not relationship:
        # Check if primary carer is emergency contact
        is_primary_carer = csv_data.get('Is the primary carer also the emergency contact for the participant?', '').strip()
        is_primary_carer_clean = is_primary_carer.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip().lower()
        
        if 'yes' in is_primary_carer_clean:
            relationship = csv_data.get('Relationship to client (Primary carer)', '').strip()
    
    return relationship if relationship else ''

def get_client_phone_numbers(csv_data):
    """Get client phone numbers (Home phone + Mobile phone + Work phone)"""
    phones = []
    
    home_phone = csv_data.get('Home phone (Contact Details of the Client)', '').strip()
    mobile_phone = csv_data.get('Mobile phone (Contact Details of the Client)', '').strip()
    work_phone = csv_data.get('Work phone (Contact Details of the Client)', '').strip()
    
    if home_phone:
        phones.append(home_phone)
    if mobile_phone:
        phones.append(mobile_phone)
    if work_phone:
        phones.append(work_phone)
    
    return '; '.join(phones) if phones else ''

def create_emergency_disaster_plan_from_data(csv_data, output_path, contact_name=None):
    """
    Create an Emergency & Disaster Plan PDF from provided data dictionary.
    
    Args:
        csv_data: Dictionary containing form data
        output_path: Path where the PDF should be saved
        contact_name: Optional name to use for Team member lookup
    """
    # Get team value to determine which active users CSV to use
    team_value = csv_data.get('Neighbourhood Care representative team', '')
    # Clean up checkbox characters
    team_value = team_value.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
    
    # Load active users based on team
    active_users = load_active_users(team_value)
    
    # Get team member name (similar to Key Contact lookup)
    team_member_name_to_use = contact_name or csv_data.get('Respondent', '')
    team_member_data = lookup_user_data(active_users, team_member_name_to_use) if team_member_name_to_use else {'name': '', 'mobile': '', 'email': ''}
    team_member_name = team_member_name_to_use if team_member_name_to_use else team_member_data.get('name', '')
    
    # Create PDF document
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    story = []
    styles = getSampleStyleSheet()
    
    # Create custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontSize=18,
        textColor=BLUE_COLOR,
        alignment=TA_LEFT,
        spaceAfter=12,
        leftIndent=0
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=BLUE_COLOR,
        alignment=TA_LEFT,
        spaceAfter=8,
        spaceBefore=12,
        leftIndent=0
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=11,
        alignment=TA_LEFT,
        spaceAfter=6,
        leading=14,
        leftIndent=0
    )
    
    table_text_style = ParagraphStyle(
        'TableText',
        parent=styles['Normal'],
        fontSize=10,
        alignment=TA_LEFT,
        spaceAfter=0,
        leading=12,
        leftIndent=0
    )
    
    # Title
    story.append(Paragraph("Emergency & Disaster Plan", title_style))
    story.append(Spacer(1, 0.2*inch))
    
    # General Information table
    first_name = csv_data.get('First name (Details of the Client)', '').strip()
    surname = csv_data.get('Surname (Details of the Client)', '').strip()
    client_name = ' '.join([p for p in [first_name, surname] if p]).strip() or ''
    client_phone = get_client_phone_numbers(csv_data)
    
    general_info_data = [
        [Paragraph("General Information", ParagraphStyle('TableHeading', parent=table_text_style, fontSize=12, textColor=colors.white, alignment=TA_CENTER))],
        ['Your name', Paragraph(client_name, table_text_style)],
        ['Your phone number', Paragraph(client_phone, table_text_style)]
    ]
    
    general_info_table = Table(general_info_data, colWidths=[2.5*inch, 3.5*inch])
    general_info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLUE_COLOR),  # Heading row
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 1), (1, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (1, 1), (1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('SPAN', (0, 0), (-1, 0))  # Span heading across all columns
    ]))
    
    story.append(general_info_table)
    story.append(Spacer(1, 0.2*inch))
    
    # Key Emergency Contacts table
    emergency_first = csv_data.get('First name (Emergency contact)', '').strip()
    emergency_surname = csv_data.get('Surname (Emergency contact)', '').strip()
    
    # Check if primary carer is emergency contact
    is_primary_carer = csv_data.get('Is the primary carer also the emergency contact for the participant?', '').strip()
    is_primary_carer_clean = is_primary_carer.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip().lower()
    
    if 'yes' in is_primary_carer_clean:
        emergency_first = csv_data.get('First name (Primary carer)', '').strip()
        emergency_surname = csv_data.get('Surname (Primary carer)', '').strip()
    
    emergency_name = ' '.join([p for p in [emergency_first, emergency_surname] if p]).strip() or ''
    emergency_phone = get_emergency_contact_phone(csv_data)
    emergency_relationship = get_emergency_contact_relationship(csv_data)
    
    emergency_contacts_data = [
        [Paragraph("Key Emergency Contacts", ParagraphStyle('TableHeading', parent=table_text_style, fontSize=12, textColor=colors.white, alignment=TA_CENTER))],
        [Paragraph("Name", ParagraphStyle('TableHeader', parent=table_text_style, fontSize=10, textColor=colors.white, alignment=TA_CENTER)),
         Paragraph("Phone", ParagraphStyle('TableHeader', parent=table_text_style, fontSize=10, textColor=colors.white, alignment=TA_CENTER)),
         Paragraph("Relationship", ParagraphStyle('TableHeader', parent=table_text_style, fontSize=10, textColor=colors.white, alignment=TA_CENTER))],
        [Paragraph(emergency_name, table_text_style), emergency_phone, Paragraph(emergency_relationship, table_text_style)]
    ]
    
    emergency_contacts_table = Table(emergency_contacts_data, colWidths=[2*inch, 2*inch, 2*inch])
    emergency_contacts_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLUE_COLOR),  # Heading row
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, 1), BLUE_COLOR),  # Header row
        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),
        ('BACKGROUND', (0, 2), (-1, 2), colors.white),  # Data row
        ('TEXTCOLOR', (0, 2), (-1, 2), colors.black),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),
        ('ALIGN', (0, 2), (-1, 2), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('SPAN', (0, 0), (-1, 0))  # Span heading across all columns
    ]))
    
    story.append(emergency_contacts_table)
    story.append(Spacer(1, 0.2*inch))
    
    # My Important Contacts table
    important_contacts_data = [
        [Paragraph("My Important Contacts", ParagraphStyle('TableHeading', parent=table_text_style, fontSize=12, textColor=colors.white, alignment=TA_CENTER))],
        ['Advocate', ''],
        ['Power of Attorney/Guardian', ''],
        ['Solicitor', ''],
        ['Insurer (home)', ''],
        ['Insurer (vehicle)', ''],
        ['Childcare/School Contact', ''],
        ['Workplace/Volunteer Contact', ''],
        ['Doctor', ''],
        ['Specialist Practitioner', ''],
        ['Private Health Cover', '']
    ]
    
    important_contacts_table = Table(important_contacts_data, colWidths=[2.5*inch, 3.5*inch])
    important_contacts_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLUE_COLOR),  # Heading row
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 1), (1, -1), 'LEFT'),
        ('FONTNAME', (0, 1), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (1, 1), (1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('SPAN', (0, 0), (-1, 0))  # Span heading across all columns
    ]))
    
    story.append(important_contacts_table)
    story.append(Spacer(1, 0.2*inch))
    
    # What are the main risks in your community table
    risks = [
        'Heatwave', 'Storm', 'Cyclone', 'Bushfire', 'Flood', 'Earthquake',
        'Landslide', 'Tsunami', 'Assault', 'Power outage', 'Gas outage',
        'Health emergency', 'House fire', 'Burglary/break-in'
    ]
    
    risks_data = []
    for risk in risks:
        risks_data.append([risk, '[ ]'])  # Use simple brackets instead of checkbox character
    
    risks_table = Table(risks_data, colWidths=[2.5*inch, 3.5*inch])
    risks_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    
    story.append(Paragraph("What are the main risks in your community?", heading_style))
    story.append(risks_table)
    story.append(Spacer(1, 0.2*inch))
    
    # How would the emergency affect you? table
    emergency_types = [
        'Heatwave', 'Storm', 'Cyclone', 'Bushfire', 'Flood', 'Earthquake',
        'Landslide', 'Tsunami', 'Assault', 'Power outage', 'Gas outage',
        'Health emergency', 'House fire', 'Burglary/break-in'
    ]
    
    emergency_affect_data = [
        [Paragraph("Emergency Type", ParagraphStyle('TableHeader', parent=table_text_style, fontSize=10, textColor=colors.white, alignment=TA_CENTER)),
         Paragraph("How you're affected", ParagraphStyle('TableHeader', parent=table_text_style, fontSize=10, textColor=colors.white, alignment=TA_CENTER))]
    ]
    
    # Add rows for each emergency type
    for emergency_type in emergency_types:
        emergency_affect_data.append([Paragraph(emergency_type, table_text_style), ''])
    
    emergency_affect_table = Table(emergency_affect_data, colWidths=[2.5*inch, 3.5*inch])
    emergency_affect_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLUE_COLOR),  # Header row
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 1), (1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    
    story.append(Paragraph("How would the emergency affect you?", heading_style))
    story.append(emergency_affect_table)
    story.append(Spacer(1, 0.2*inch))
    
    # Complete all applicable sections table
    sections = [
        'My Emergency & Disaster Plan',
        'Communication',
        'Management of Health',
        'Assistive Technology (AT)',
        'Personal Support',
        'Assistance animals and pets',
        'Transportation',
        'Living Situation',
        'Social Connectedness',
        'Other'
    ]
    
    sections_data = []
    for i, section in enumerate(sections):
        if i == 0:  # First section gets blue background and white text
            sections_data.append([Paragraph(section, ParagraphStyle('TableTextBlue', parent=table_text_style, fontSize=10, textColor=colors.white)), ''])
        else:
            sections_data.append([section, ''])
    
    sections_table = Table(sections_data, colWidths=[2.5*inch, 3.5*inch])
    sections_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), BLUE_COLOR),  # First row first column
        ('TEXTCOLOR', (0, 0), (0, 0), colors.white),
        ('BACKGROUND', (1, 0), (1, 0), colors.white),  # First row second column
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),  # Rest of rows
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    
    story.append(Paragraph("Complete all applicable sections & if not applicable, mark as \"N/A\"", heading_style))
    story.append(sections_table)
    story.append(Spacer(1, 0.2*inch))
    
    # Final table with signatures (4 columns, 3 rows)
    # Get modification date - try to get from PDF file modification time or use current date
    from datetime import datetime
    mod_date = datetime.now().strftime('%d/%m/%Y')
    
    final_data = [
        ['Client\'s Name', Paragraph(client_name, table_text_style), 'Team member\'s name', Paragraph(team_member_name, table_text_style)],
        ['Signature', '', 'Signature', ''],
        ['Date', mod_date, 'Date', mod_date]
    ]
    
    final_table = Table(final_data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.5*inch])
    final_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'LEFT'),
        ('ALIGN', (3, 0), (3, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
        ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
        ('FONTNAME', (3, 0), (3, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))
    
    story.append(final_table)
    
    # Build PDF
    doc.build(story)
    print("Emergency & Disaster Plan PDF created successfully!")

if __name__ == "__main__":
    create_service_agreement()

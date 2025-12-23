import os
import csv
import zipfile
from pathlib import Path
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import tempfile
import uuid
from typing import Dict, List

try:
    import pdfplumber  # text extraction
except Exception:
    pdfplumber = None
try:
    from pypdf import PdfReader  # form fields
except Exception:
    PdfReader = None

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'your-secret-key-change-in-production')

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Create upload directory if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Handle Render's src directory structure - add parent directory to path if needed
import sys
current_dir = os.getcwd()
# If we're in a src directory, add parent to path so we can import from root
if current_dir.endswith('/src') or current_dir.endswith('\\src'):
    parent_dir = os.path.dirname(current_dir)
    if parent_dir and parent_dir not in sys.path:
        sys.path.insert(0, parent_dir)
    # Also ensure current directory is in path
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)
else:
    # Ensure current directory is in path
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)

OUTPUT_FIELDS = [
    "Salutation",
    "First Name",
    "Middle Name",
    "Family Name",
    "Display Name",
    "Date of Birth",
    "Gender",
    "Address",
    "Address Unit/Apartment Number",
    "General Information",
    "Phone Number",
    "Mobile Number",
    "Email",
    "Marital Status",
    "Nationality",
    "Languages",
    "NDIS Number",
    "Age Care Recipient ID",
    "Reference Number",
    "Purchase Order Number",
]

HEADER_VARIANTS: Dict[str, List[str]] = {
    # Details of the Client
    "ndis number": ["ndis number", "ndis number (details of the client)"],
    "first name": ["first name", "first name (details of the client)"],
    "middle name": ["middle name", "middle name (details of the client)"],
    "surname": ["surname", "surname (details of the client)"],
    "date of birth": ["date of birth", "date of birth (details of the client)"],
    "gender": ["gender", "gender (details of the client)"],
    # Contact Details of the Client
    "home address": ["home address", "home address (contact details of the client)"],
    "home phone": ["home phone", "home phone (contact details of the client)"],
    "work phone": ["work phone", "work phone (contact details of the client)"],
    "mobile phone": ["mobile phone", "mobile phone (contact details of the client)"],
    "email address": ["email address", "email address (contact details of the client)"],
}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def normalize_key(key: str) -> str:
    return str(key or "").strip().lower()


def build_normalized_row(row: dict) -> dict:
    return {normalize_key(k): v for k, v in row.items()}


def get_value_from_normalized_row(row_norm: dict, *possible_headers: str) -> str:
    for header in possible_headers:
        key = normalize_key(header)
        # direct
        if key in row_norm and row_norm[key] not in (None, ""):
            return str(row_norm[key]).strip()
        # try variants map
        if key in HEADER_VARIANTS:
            for variant in HEADER_VARIANTS[key]:
                vkey = normalize_key(variant)
                if vkey in row_norm and row_norm[vkey] not in (None, ""):
                    return str(row_norm[vkey]).strip()
    return ""

def is_valid_phone_number(value: str) -> bool:
    """Check if a value looks like a phone number"""
    if not value:
        return False
    
    # Remove common phone number formatting characters
    cleaned = ''.join(c for c in value if c.isdigit() or c in ['+', '-', '(', ')', ' ', '.', 'x', 'X'])
    
    # Must contain digits
    digits_only = ''.join(c for c in cleaned if c.isdigit())
    if len(digits_only) < 6:  # Too short to be a phone number
        return False
    if len(digits_only) > 20:  # Too long to be a phone number
        return False
    
    # Must NOT contain @ (email addresses)
    if '@' in value:
        return False
    
    # Should not be all letters (check if it's mostly letters)
    letters = ''.join(c for c in value if c.isalpha())
    if len(letters) > len(digits_only):
        return False
    
    return True

def format_date_dd_mm_yyyy(date_str):
    """
    Convert date string to DD/MM/YYYY format.
    Handles various input formats: YYYY-MM-DD, MM/DD/YYYY, DD/MM/YYYY, DD-MM-YYYY, etc.
    """
    if not date_str:
        return ""
    
    date_str = date_str.strip()
    if not date_str:
        return ""
    
    import re
    from datetime import datetime
    
    # Try to parse common date formats
    date_formats = [
        '%Y-%m-%d',      # 2023-12-25
        '%d/%m/%Y',      # 25/12/2023
        '%m/%d/%Y',      # 12/25/2023
        '%d-%m-%Y',      # 25-12-2023
        '%Y/%m/%d',      # 2023/12/25
        '%d.%m.%Y',      # 25.12.2023
        '%d %m %Y',      # 25 12 2023
    ]
    
    for fmt in date_formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            # Return in DD/MM/YYYY format
            return dt.strftime('%d/%m/%Y')
        except ValueError:
            continue
    
    # If no format matched, try to extract numbers and rearrange
    # Look for pattern like: 3 numbers separated by /, -, or space
    try:
        numbers = re.findall(r'\d+', date_str)
        if len(numbers) >= 3:
            # If first number is 4 digits, assume YYYY-MM-DD or YYYY/MM/DD
            if len(numbers[0]) == 4:
                year, month, day = numbers[0], numbers[1], numbers[2]
            # If last number is 4 digits, assume DD/MM/YYYY or MM/DD/YYYY
            elif len(numbers[-1]) == 4:
                # Try to determine if it's DD/MM/YYYY or MM/DD/YYYY
                # If middle number > 12, it must be DD/MM/YYYY
                try:
                    if int(numbers[1]) > 12:
                        day, month, year = numbers[0], numbers[1], numbers[2]
                    else:
                        # Ambiguous - assume DD/MM/YYYY (day first)
                        day, month, year = numbers[0], numbers[1], numbers[2]
                except (ValueError, IndexError):
                    # If conversion fails, default to DD/MM/YYYY
                    day, month, year = numbers[0], numbers[1], numbers[2]
            else:
                # Default to DD/MM/YYYY
                day, month, year = numbers[0], numbers[1], numbers[2]
            
            # Format with leading zeros
            day = day.zfill(2)
            month = month.zfill(2)
            return f"{day}/{month}/{year}"
    except (ValueError, IndexError, AttributeError):
        # If anything goes wrong, return original string
        pass
    
    # If we can't parse it, return as-is
    return date_str


def build_output_row(row):
    """
    Build output row with correct field mappings:
    - Details of the Client:
      - 'NDIS number' → 'NDIS Number'
      - 'First name' → 'First Name'
      - 'Middle name' → 'Middle Name'
      - 'Surname' → 'Family Name'
      - 'Display Name' = 'First name' + 'Surname'
      - 'Date of birth' → 'Date of Birth' (formatted as DD/MM/YYYY)
      - 'Gender' → 'Gender'
    - Contact Details of the Client:
      - 'Home address' → 'Address'
      - 'Home phone' and/or 'Work phone' (separated by ';') → 'Phone Number'
      - 'Mobile phone' → 'Mobile Number'
      - 'Email address' → 'Email'
    """
    row_norm = build_normalized_row(row)
    
    # Details of the Client
    first_name = get_value_from_normalized_row(row_norm, "first name")
    middle_name = get_value_from_normalized_row(row_norm, "middle name")
    surname = get_value_from_normalized_row(row_norm, "surname")
    ndis_number = get_value_from_normalized_row(row_norm, "ndis number")
    dob_raw = get_value_from_normalized_row(row_norm, "date of birth")
    dob = format_date_dd_mm_yyyy(dob_raw)  # Format as DD/MM/YYYY
    gender = get_value_from_normalized_row(row_norm, "gender")
    
    # Contact Details of the Client
    home_address = get_value_from_normalized_row(row_norm, "home address")
    home_phone = get_value_from_normalized_row(row_norm, "home phone")
    work_phone = get_value_from_normalized_row(row_norm, "work phone")
    mobile_phone = get_value_from_normalized_row(row_norm, "mobile phone")
    email_address = get_value_from_normalized_row(row_norm, "email address")
    # For client export, only use the first email if multiple are present
    if email_address and ';' in email_address:
        email_address = email_address.split(';')[0].strip()
    
    # Build display name from first name + surname
    display_name = " ".join([p for p in [first_name, surname] if p]).strip()
    
    # Use only home phone (work phone is not included)
    phone_number = home_phone if home_phone else ""

    return {
        "Salutation": "They",
        "First Name": first_name,  # 'First name' (Details of the Client) → 'First Name'
        "Middle Name": middle_name,  # 'Middle name' (Details of the Client) → 'Middle Name'
        "Family Name": surname,  # 'Surname' (Details of the Client) → 'Family Name'
        "Display Name": display_name,  # 'First name' + 'Surname'
        "Date of Birth": dob,  # 'Date of birth' (Details of the Client) → 'Date of Birth'
        "Gender": gender,  # 'Gender' (Details of the Client) → 'Gender'
        "Address": home_address,  # 'Home address' (Contact Details of the Client) → 'Address'
        "Address Unit/Apartment Number": "",
        "General Information": "",
        "Phone Number": phone_number,  # 'Home phone' and/or 'Work phone' (separated by ';') (Contact Details of the Client) → 'Phone Number'
        "Mobile Number": mobile_phone if mobile_phone else "",  # 'Mobile phone' (Contact Details of the Client) → 'Mobile Number'
        "Email": email_address,  # 'Email address' (Contact Details of the Client) → 'Email'
        "Marital Status": "",
        "Nationality": "",
        "Languages": "",
        "NDIS Number": ndis_number,  # 'NDIS number' (Details of the Client) → 'NDIS Number'
        "Age Care Recipient ID": "",
        "Reference Number": "",
        "Purchase Order Number": "",
    }

def transform_csv(input_path, output_path):
    with open(input_path, 'r', newline='', encoding='utf-8') as f_in:
        reader = csv.DictReader(f_in)
        rows = list(reader)

    output_rows = []
    for row in rows:
        output_rows.append(build_output_row(row))

    with open(output_path, 'w', newline='', encoding='utf-8') as f_out:
        writer = csv.DictWriter(f_out, fieldnames=OUTPUT_FIELDS)
        writer.writeheader()
        for out_row in output_rows:
            writer.writerow(out_row)


def extract_pdf_fields_pdfreader(pdf_path: str) -> dict:
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


def parse_pdf_to_row(pdf_path: str) -> dict:
    # Strategy: prefer form fields; fallback to text scanning
    data: Dict[str, str] = {}
    fields = extract_pdf_fields_pdfreader(pdf_path)
    
    # Check if we actually got any non-empty values from form fields
    has_form_data = False
    if fields:
        # try map common field names by fuzzy containment
        def find_in_fields(*candidates):
            cand_norm = [normalize_key(c) for c in candidates]
            for key, val in fields.items():
                k = normalize_key(key)
                for c in cand_norm:
                    if c in k:
                        value = str(val).strip()
                        if value:  # Only return non-empty values
                            return value
            return ""

        data["first name"] = find_in_fields("first name")
        data["middle name"] = find_in_fields("middle name")
        data["surname"] = find_in_fields("surname", "family name", "last name")
        data["ndis number"] = find_in_fields("ndis")
        data["date of birth"] = find_in_fields("date of birth", "dob")
        data["gender"] = find_in_fields("gender")
        data["home address"] = find_in_fields("home address", "address")
        data["home phone"] = find_in_fields("home phone")
        data["work phone"] = find_in_fields("work phone")
        data["mobile phone"] = find_in_fields("mobile")
        data["email address"] = find_in_fields("email")
        
        # Check if we got any actual data
        has_form_data = any(data.values())
    
    # If form fields didn't yield data, try text extraction
    if not has_form_data:
        # Fallback to text scanning - ONLY extract from correct sections
        text = extract_pdf_text_pdfplumber(pdf_path)
        if text:
            lines = [l.strip() for l in text.splitlines() if l.strip()]
            
            # Identify section boundaries
            in_details_section = False
            in_contact_section = False
            section_starts = []
            
            for i, line in enumerate(lines):
                line_lower = normalize_key(line)
                if "details of the client" in line_lower and "contact" not in line_lower:
                    in_details_section = True
                    in_contact_section = False
                    section_starts.append(("details", i))
                elif "contact details of the client" in line_lower:
                    in_details_section = False
                    in_contact_section = True
                    section_starts.append(("contact", i))
                elif any(x in line_lower for x in ["needs of the client", "ndis information", "support items", "formal supports", 
                                                   "primary carer", "important people", "home life", "health information", 
                                                   "care requirements", "behaviour requirements", "other information", "consents"]):
                    in_details_section = False
                    in_contact_section = False

            def find_value_in_section(label_patterns: List[str], section_type: str, exclude_email: bool = False) -> str:
                """Find value only in the specified section (details or contact)"""
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
                            section_end = len(lines)
                        break
                
                if section_start is None:
                    return ""
                
                # Only search within this section
                for i in range(section_start, section_end):
                    line = lines[i]
                    line_lower = normalize_key(line)
                    
                    for pattern in label_patterns:
                        pattern_lower = normalize_key(pattern)
                        
                        # Must match pattern exactly, not as part of another phrase
                        # e.g., "first name" should match "First name" but not "Formal support (1) first name"
                        if pattern_lower == line_lower or (pattern_lower in line_lower and not any(x in line_lower for x in 
                            ["formal support", "informal support", "primary carer", "emergency contact", "plan manager"])):
                            
                            # Look ONLY at the immediate next line for the value
                            if i + 1 < section_end:
                                next_line = lines[i + 1].strip()
                                if next_line:
                                    next_line_lower = normalize_key(next_line)
                                    
                                    # Check if the next line is another field label by comparing with HEADER_VARIANTS
                                    is_field_label = False
                                    for other_pattern_list in HEADER_VARIANTS.values():
                                        for other_pattern in other_pattern_list:
                                            other_pattern_lower = normalize_key(other_pattern)
                                            # Exact match means it's a label
                                            if other_pattern_lower == next_line_lower:
                                                is_field_label = True
                                                break
                                        if is_field_label:
                                            break
                                    
                                    # If it's not a field label and doesn't end with ), it's a value
                                    if not is_field_label and not next_line.endswith(')'):
                                        # If looking for phone, exclude email addresses
                                        if exclude_email and '@' in next_line:
                                            return ""
                                        # This is the value
                                        return next_line
                
                return ""
            
            # Extract from Details of the Client section
            data["first name"] = find_value_in_section(HEADER_VARIANTS["first name"], "details")
            data["middle name"] = find_value_in_section(HEADER_VARIANTS["middle name"], "details")
            data["surname"] = find_value_in_section(HEADER_VARIANTS["surname"], "details")
            data["ndis number"] = find_value_in_section(HEADER_VARIANTS["ndis number"], "details")
            data["date of birth"] = find_value_in_section(HEADER_VARIANTS["date of birth"], "details")
            data["gender"] = find_value_in_section(HEADER_VARIANTS["gender"], "details")
            
            # Extract from Contact Details of the Client section
            data["home address"] = find_value_in_section(HEADER_VARIANTS["home address"], "contact")
            data["home phone"] = find_value_in_section(HEADER_VARIANTS["home phone"], "contact", exclude_email=True)
            data["work phone"] = find_value_in_section(HEADER_VARIANTS["work phone"], "contact", exclude_email=True)
            data["mobile phone"] = find_value_in_section(HEADER_VARIANTS["mobile phone"], "contact", exclude_email=True)
            data["email address"] = find_value_in_section(HEADER_VARIANTS["email address"], "contact")
    
    return data

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/debug')
def debug_info():
    """Diagnostic endpoint to help troubleshoot Render deployment issues"""
    import sys
    import traceback
    info = {
        'current_directory': os.getcwd(),
        'python_version': sys.version,
        'python_path': sys.path[:10],
        'files_in_current_dir': [f for f in os.listdir('.') if f.endswith('.py')][:20],
        'create_final_tables_exists': os.path.exists('create_final_tables.py'),
        'app_py_exists': os.path.exists('app.py'),
    }
    
    # Try to import the module
    try:
        import create_final_tables
        info['module_imported'] = True
        info['module_file'] = getattr(create_final_tables, '__file__', 'unknown')
        info['module_has_parse_pdf_to_data'] = hasattr(create_final_tables, 'parse_pdf_to_data')
        info['module_has_load_ndis'] = hasattr(create_final_tables, 'load_ndis_support_items')
        info['module_has_load_users'] = hasattr(create_final_tables, 'load_active_users')
        info['module_attributes'] = [a for a in dir(create_final_tables) if not a.startswith('_')][:30]
    except Exception as e:
        info['module_imported'] = False
        info['module_import_error'] = str(e)
        info['module_import_traceback'] = traceback.format_exc()
    
    # Format as HTML for easy reading
    html = "<h1>Render Deployment Diagnostics</h1><pre>"
    html += "\n".join([f"{k}: {v}" for k, v in info.items()])
    html += "</pre>"
    html += "<p><a href='/'>Back to main page</a></p>"
    return html

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(url_for('index'))
    
    file = request.files['file']
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    # Check what to generate
    generate_csv = request.form.get('generate_csv') == '1'
    generate_service_agreement = request.form.get('generate_service_agreement') == '1'
    generate_emergency_plan = request.form.get('generate_emergency_plan') == '1'
    generate_service_estimate = request.form.get('generate_service_estimate') == '1'
    generate_risk_assessment = request.form.get('generate_risk_assessment') == '1'
    generate_support_plan = request.form.get('generate_support_plan') == '1'
    generate_medication_plan = request.form.get('generate_medication_plan') == '1'
    
    if not generate_csv and not generate_service_agreement and not generate_emergency_plan and not generate_service_estimate and not generate_risk_assessment and not generate_support_plan and not generate_medication_plan:
        flash('Please select at least one output to generate')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        # Add unique identifier to avoid conflicts
        unique_filename = f"{uuid.uuid4()}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        
        # Ensure upload directory exists
        os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
        
        try:
            file.save(filepath)
        except Exception as save_error:
            flash(f'Error saving file: {str(save_error)}')
            return redirect(url_for('index'))
        
        try:
            # Parse PDF data using the working function from create_final_tables
            # Try multiple import strategies to handle different Render configurations
            parse_pdf_to_data = None
            load_ndis_support_items = None
            load_active_users = None
            
            import sys
            import traceback
            
            # Strategy 1: Normal import
            import_error_details = None
            module_execution_error = None
            
            # First, try to find and load the file directly
            current_dir = os.getcwd()
            file_path = os.path.join(current_dir, 'create_final_tables.py')
            parent_dir = os.path.dirname(current_dir)
            parent_file_path = os.path.join(parent_dir, 'create_final_tables.py') if parent_dir else None
            
            # Try direct file import first (most reliable)
            if os.path.exists(file_path):
                try:
                    import importlib.util
                    if 'create_final_tables' in sys.modules:
                        del sys.modules['create_final_tables']
                    
                    spec = importlib.util.spec_from_file_location("create_final_tables", file_path)
                    if spec and spec.loader:
                        create_final_tables_module = importlib.util.module_from_spec(spec)
                        spec.loader.exec_module(create_final_tables_module)
                        
                        # Now try to get the functions
                        parse_pdf_to_data = getattr(create_final_tables_module, 'parse_pdf_to_data', None)
                        load_ndis_support_items = getattr(create_final_tables_module, 'load_ndis_support_items', None)
                        load_active_users = getattr(create_final_tables_module, 'load_active_users', None)
                        
                        if parse_pdf_to_data and load_ndis_support_items and load_active_users:
                            print("✓ Successfully loaded functions via direct file import")
                            # Store in sys.modules for future imports
                            sys.modules['create_final_tables'] = create_final_tables_module
                        else:
                            print(f"⚠ Direct import loaded module but functions missing: parse={parse_pdf_to_data is not None}, ndis={load_ndis_support_items is not None}, users={load_active_users is not None}")
                except Exception as direct_import_err:
                    print(f"Direct file import failed: {direct_import_err}")
                    import traceback
                    print(traceback.format_exc())
            
            # If direct import didn't work, try normal import
            if not parse_pdf_to_data or not load_ndis_support_items or not load_active_users:
                try:
                    from create_final_tables import parse_pdf_to_data, load_ndis_support_items, load_active_users
                    print("✓ Successfully loaded functions via normal import")
                except (ImportError, AttributeError) as e:
                    import_error_details = str(e)
                    import traceback
                    import_error_details += f"\n{traceback.format_exc()}"
                    print(f"Normal import failed: {import_error_details}")
                
                # Strategy 2: Import module first, then get attributes
                try:
                    # Try importing from current directory
                    print("Trying Strategy 2: Import module then get attributes...")
                    import create_final_tables
                    print(f"Module imported. File: {getattr(create_final_tables, '__file__', 'unknown')}")
                    print(f"Module dir: {dir(create_final_tables)[:10]}")
                    
                    # Try to access functions - this might fail if module execution was incomplete
                    try:
                        parse_pdf_to_data = getattr(create_final_tables, 'parse_pdf_to_data', None)
                        load_ndis_support_items = getattr(create_final_tables, 'load_ndis_support_items', None)
                        load_active_users = getattr(create_final_tables, 'load_active_users', None)
                        print(f"After getattr - parse_pdf_to_data: {parse_pdf_to_data is not None}")
                    except Exception as attr_err:
                        module_execution_error = f"Error accessing attributes: {attr_err}"
                        import traceback
                        module_execution_error += f"\n{traceback.format_exc()}"
                        print(f"Error accessing attributes: {module_execution_error}")
                except Exception:
                    # Strategy 3: Try from parent directory (if we're in src/)
                    try:
                        current_dir = os.getcwd()
                        parent_dir = os.path.dirname(current_dir)
                        if parent_dir and parent_dir not in sys.path:
                            sys.path.insert(0, parent_dir)
                        
                        # Reload module from parent
                        import importlib
                        if 'create_final_tables' in sys.modules:
                            del sys.modules['create_final_tables']
                        import create_final_tables
                        parse_pdf_to_data = getattr(create_final_tables, 'parse_pdf_to_data', None)
                        load_ndis_support_items = getattr(create_final_tables, 'load_ndis_support_items', None)
                        load_active_users = getattr(create_final_tables, 'load_active_users', None)
                    except Exception:
                        pass
                
                # Strategy 4: Try direct file import if module import worked but functions missing
                if not parse_pdf_to_data:
                    try:
                        current_dir = os.getcwd()
                        file_path = os.path.join(current_dir, 'create_final_tables.py')
                        # Check if file exists in current directory
                        if os.path.exists(file_path):
                            import importlib.util
                            # Remove from sys.modules if already loaded (might be broken)
                            if 'create_final_tables' in sys.modules:
                                del sys.modules['create_final_tables']
                            
                            spec = importlib.util.spec_from_file_location(
                                "create_final_tables", 
                                file_path
                            )
                            if spec and spec.loader:
                                create_final_tables = importlib.util.module_from_spec(spec)
                                # Execute the module - this will show any errors
                                try:
                                    spec.loader.exec_module(create_final_tables)
                                    parse_pdf_to_data = getattr(create_final_tables, 'parse_pdf_to_data', None)
                                    load_ndis_support_items = getattr(create_final_tables, 'load_ndis_support_items', None)
                                    load_active_users = getattr(create_final_tables, 'load_active_users', None)
                                except Exception as exec_err:
                                    print(f"ERROR executing create_final_tables module: {exec_err}")
                                    import traceback
                                    print(traceback.format_exc())
                                    raise
                    except Exception as file_import_err:
                        print(f"ERROR in direct file import: {file_import_err}")
                        import traceback
                        print(traceback.format_exc())
            
            # If still not found, provide detailed diagnostics
            if not parse_pdf_to_data or not load_ndis_support_items or not load_active_users:
                current_dir = os.getcwd()
                files_in_dir = []
                try:
                    files_in_dir = [f for f in os.listdir('.') if f.endswith('.py')][:10]
                except Exception:
                    pass
                
                # Check if file exists in current or parent directory
                file_in_current = os.path.exists(os.path.join(current_dir, 'create_final_tables.py'))
                parent_dir = os.path.dirname(current_dir)
                file_in_parent = os.path.exists(os.path.join(parent_dir, 'create_final_tables.py')) if parent_dir else False
                
                python_path = ', '.join(sys.path[:5])
                
                # Try to see what's actually in the module and diagnose the issue
                module_contents = []
                module_file_path = None
                module_load_error = None
                module_compile_error = None
                try:
                    # First, try to compile the file to check for syntax errors
                    file_path = os.path.join(current_dir, 'create_final_tables.py')
                    if os.path.exists(file_path):
                        try:
                            with open(file_path, 'r', encoding='utf-8') as f:
                                code = f.read()
                            compile(code, file_path, 'exec')
                            print("✓ File compiles successfully (no syntax errors)")
                        except SyntaxError as syn_err:
                            module_compile_error = f"Syntax error in file: {syn_err}"
                            import traceback
                            module_compile_error += f"\n{traceback.format_exc()}"
                            print(f"✗ Syntax error: {module_compile_error}")
                        except Exception as compile_err:
                            module_compile_error = f"Error compiling file: {compile_err}"
                            print(f"✗ Compile error: {module_compile_error}")
                    
                    # Try to reload the module fresh
                    if 'create_final_tables' in sys.modules:
                        del sys.modules['create_final_tables']
                        print("Removed module from cache")
                    
                    print("Attempting to import module...")
                    # Try importing with explicit error handling
                    try:
                        import create_final_tables
                        module_file_path = getattr(create_final_tables, '__file__', 'unknown')
                        print(f"✓ Module imported successfully. File: {module_file_path}")
                    except Exception as import_ex:
                        print(f"✗ Module import failed: {import_ex}")
                        import traceback
                        print(traceback.format_exc())
                        raise
                    
                    # Get all attributes
                    all_attrs = dir(create_final_tables)
                    module_contents = [attr for attr in all_attrs if not attr.startswith('_')][:30]
                    print(f"Module has {len(all_attrs)} total attributes, {len(module_contents)} public attributes")
                    print(f"Public attributes: {module_contents}")
                    
                    # Check if functions exist but are None or not callable
                    funcs_to_check = ['parse_pdf_to_data', 'load_ndis_support_items', 'load_active_users']
                    for func_name in funcs_to_check:
                        if hasattr(create_final_tables, func_name):
                            func = getattr(create_final_tables, func_name)
                            print(f"✓ {func_name} found - type: {type(func)}, callable: {callable(func) if func else 'N/A'}")
                        else:
                            print(f"✗ {func_name} NOT found in module")
                    
                    # Try to search for the function definition in the file
                    print("\nSearching file for function definitions...")
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            file_content = f.read()
                            for func_name in funcs_to_check:
                                if f'def {func_name}' in file_content:
                                    print(f"✓ Found 'def {func_name}' in file content")
                                else:
                                    print(f"✗ 'def {func_name}' NOT found in file content")
                    except Exception as file_read_err:
                        print(f"Could not read file to check function definitions: {file_read_err}")
                        
                except Exception as module_err:
                    module_load_error = str(module_err)
                    import traceback
                    module_load_error += f"\n{traceback.format_exc()}"
                    print(f"✗ ERROR loading module: {module_load_error}")
                
                # Try to read the file and check for syntax errors
                file_content_check = "Not checked"
                try:
                    file_path = os.path.join(current_dir, 'create_final_tables.py')
                    if os.path.exists(file_path):
                        # Check file size
                        file_size = os.path.getsize(file_path)
                        # Read first and last few lines to verify it's a valid Python file
                        with open(file_path, 'r', encoding='utf-8') as f:
                            first_lines = [f.readline() for _ in range(5)]
                            f.seek(max(0, file_size - 500))  # Last 500 bytes
                            last_lines = f.readlines()[-5:]
                        file_content_check = f"File size: {file_size} bytes, First line: {first_lines[0][:50] if first_lines else 'empty'}, Has 'def parse_pdf_to_data': {'def parse_pdf_to_data' in ''.join(first_lines + last_lines)}"
                except Exception as file_check_err:
                    file_content_check = f"Error reading file: {file_check_err}"
                
                diagnostic_msg = (
                    f"CRITICAL: Cannot import required functions from create_final_tables\n"
                    f"Current directory: {current_dir}\n"
                    f"Parent directory: {parent_dir}\n"
                    f"File exists in current dir: {file_in_current}\n"
                    f"File exists in parent dir: {file_in_parent}\n"
                    f"Python files in current directory: {', '.join(files_in_dir) if files_in_dir else 'None found'}\n"
                    f"Python path (first 5): {python_path}\n"
                    f"Initial import error: {import_error_details or 'None'}\n"
                    f"Module execution error: {module_execution_error or 'None'}\n"
                    f"Module compile error: {module_compile_error or 'None'}\n"
                    f"Module file path: {module_file_path or 'Module not loaded'}\n"
                    f"Module load error: {module_load_error or 'None'}\n"
                    f"File content check: {file_content_check}\n"
                    f"Module attributes found: {', '.join(module_contents) if module_contents else 'Could not load module'}\n"
                    f"Functions found: parse_pdf_to_data={parse_pdf_to_data is not None}, "
                    f"load_ndis_support_items={load_ndis_support_items is not None}, "
                    f"load_active_users={load_active_users is not None}\n"
                    f"\nTROUBLESHOOTING:\n"
                    f"1. If current dir ends with '/src', check Render 'Root Directory' setting (should be '.' or empty, NOT 'src')\n"
                    f"2. Verify create_final_tables.py is in the root of your GitHub repo (not in src/ subdirectory)\n"
                    f"3. Check Render build logs AND startup logs for any errors during deployment\n"
                    f"4. Look for 'STARTUP: Verifying create_final_tables imports...' in logs to see what happened at startup\n"
                    f"5. The module may have import errors - check if all dependencies in create_final_tables.py are installed\n"
                    f"6. The module may fail during execution (e.g., trying to open CSV files that don't exist)\n"
                    f"7. Try 'Clear build cache & deploy' in Render dashboard\n"
                    f"8. Visit /debug endpoint for more diagnostic information"
                )
                # Print full diagnostics to logs
                print("=" * 80)
                print("DETAILED DIAGNOSTICS:")
                print(diagnostic_msg)
                print("=" * 80)
                
                # Also include key info in the user-facing error
                error_summary = (
                    f"Failed to import functions from create_final_tables. "
                    f"Current dir: {current_dir}. "
                    f"File exists: {file_in_current}. "
                )
                if module_load_error:
                    error_summary += f"Module load error: {module_load_error[:200]}. "
                if import_error_details:
                    error_summary += f"Import error: {import_error_details[:200]}. "
                error_summary += "Check Render logs for full details. See RENDER_FIX_STEPS.md"
                
                raise ImportError(error_summary)
            
            pdf_data = parse_pdf_to_data(filepath)
            
            # Load CSV files once (performance optimization - avoid loading multiple times)
            ndis_items = None
            active_users = None
            team_value = pdf_data.get('Neighbourhood Care representative team', '')
            # Clean up checkbox characters
            team_value = team_value.replace('\uf0d7', '').replace('•', '').replace('●', '').replace('☐', '').replace('☑', '').replace('✓', '').strip()
            
            # Pre-load CSV files if any document needs them
            if generate_service_agreement or generate_service_estimate:
                ndis_items = load_ndis_support_items()
            if generate_service_agreement or generate_emergency_plan or generate_risk_assessment or generate_support_plan or generate_medication_plan:
                active_users = load_active_users(team_value)
            
            output_files = []
            
            # Generate CSV if requested
            if generate_csv:
                # Convert pdf_data format to the format expected by build_output_row
                # For client export, only use the first email if multiple are present
                email_address = pdf_data.get('Email address (Contact Details of the Client)', '')
                if email_address and ';' in email_address:
                    email_address = email_address.split(';')[0].strip()
                
                parsed = {
                    'first name': pdf_data.get('First name (Details of the Client)', ''),
                    'middle name': pdf_data.get('Middle name (Details of the Client)', ''),
                    'surname': pdf_data.get('Surname (Details of the Client)', ''),
                    'ndis number': pdf_data.get('NDIS number (Details of the Client)', ''),
                    'date of birth': pdf_data.get('Date of birth (Details of the Client)', ''),
                    'gender': pdf_data.get('Gender (Details of the Client)', ''),
                    'home address': pdf_data.get('Home address (Contact Details of the Client)', ''),
                    'home phone': pdf_data.get('Home phone (Contact Details of the Client)', ''),
                    'work phone': pdf_data.get('Work phone (Contact Details of the Client)', ''),
                    'mobile phone': pdf_data.get('Mobile phone (Contact Details of the Client)', ''),
                    'email address': email_address
                }
                
                output_filename = f"transformed_{unique_filename}.csv"
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                with open(output_path, 'w', newline='', encoding='utf-8') as f_out:
                    writer = csv.DictWriter(f_out, fieldnames=OUTPUT_FIELDS)
                    writer.writeheader()
                    writer.writerow(build_output_row(parsed))
                output_files.append(('csv', output_path, 'client_export.csv'))
            
            # Get contact name once (used by multiple documents)
            contact_name = request.form.get('contact_name', '').strip()
            
            # Generate Service Agreement PDF if requested
            if generate_service_agreement:
                # Import the service agreement generation function
                from create_final_tables import create_service_agreement_from_data
                sa_filename = f"service_agreement_{unique_filename}.pdf"
                sa_path = os.path.join(app.config['UPLOAD_FOLDER'], sa_filename)
                # Pass source PDF path for signature extraction and pre-loaded data
                create_service_agreement_from_data(pdf_data, sa_path, contact_name, filepath, ndis_items, active_users)
                output_files.append(('pdf', sa_path, 'Service Agreement.pdf'))
            
            # Generate Emergency & Disaster Plan PDF if requested
            if generate_emergency_plan:
                # Import the emergency plan generation function
                from create_final_tables import create_emergency_disaster_plan_from_data
                edp_filename = f"emergency_disaster_plan_{unique_filename}.pdf"
                edp_path = os.path.join(app.config['UPLOAD_FOLDER'], edp_filename)
                create_emergency_disaster_plan_from_data(pdf_data, edp_path, contact_name, active_users)
                output_files.append(('pdf', edp_path, 'Emergency & Disaster Plan.pdf'))
            
            # Generate Service Estimate CSV if requested
            if generate_service_estimate:
                # Import the service estimate generation function
                from create_final_tables import create_service_estimate_csv
                se_filename = f"service_estimate_{unique_filename}.csv"
                se_path = os.path.join(app.config['UPLOAD_FOLDER'], se_filename)
                create_service_estimate_csv(pdf_data, se_path, contact_name, ndis_items)
                output_files.append(('csv', se_path, 'Service Estimate.csv'))
            
            # Generate Risk Assessment PDF if requested
            if generate_risk_assessment:
                # Import the risk assessment generation function
                from create_final_tables import create_risk_assessment_from_data
                ra_filename = f"risk_assessment_{unique_filename}.pdf"
                ra_path = os.path.join(app.config['UPLOAD_FOLDER'], ra_filename)
                create_risk_assessment_from_data(pdf_data, ra_path, contact_name, active_users)
                output_files.append(('pdf', ra_path, 'Risk Assessment.pdf'))
            
            # Generate Support Plan DOCX if requested
            if generate_support_plan:
                # Import the support plan generation function
                from create_final_tables import create_support_plan_from_data
                from datetime import datetime
                import re
                
                # Extract client information for filename
                first_name = pdf_data.get('First name (Details of the Client)', '').strip()
                surname = pdf_data.get('Surname (Details of the Client)', '').strip()
                dob_str = pdf_data.get('Date of birth (Details of the Client)', '').strip()
                ndis_number = pdf_data.get('NDIS number (Details of the Client)', '').strip()
                
                # Extract year from date of birth
                year = None
                if dob_str:
                    year_match = re.search(r'\b(19|20)\d{2}\b', dob_str)
                    if year_match:
                        year = year_match.group(0)
                
                # If no year from DOB, use current year
                if not year:
                    year = datetime.now().strftime('%Y')
                
                # Extract ID from NDIS number (first 6 digits) or generate one
                client_id = ''
                if ndis_number:
                    digits = re.sub(r'\D', '', ndis_number)
                    if len(digits) >= 6:
                        client_id = digits[:6]
                    elif len(digits) > 0:
                        client_id = digits.ljust(6, '0')
                
                # If no ID from NDIS, generate a simple ID from timestamp
                if not client_id:
                    client_id = datetime.now().strftime('%H%M%S')
                
                # Build filename: "Support Plan - [First Name] [Last Name] [Year] - [ID].docx"
                name_part = f"{first_name} {surname}".strip() if (first_name or surname) else "test test"
                sp_filename = f"Support Plan - {name_part} {year} - {client_id}.pdf"
                sp_path = os.path.join(app.config['UPLOAD_FOLDER'], sp_filename)
                create_support_plan_from_data(pdf_data, sp_path, contact_name, active_users)
                output_files.append(('pdf', sp_path, sp_filename))
            
            # Generate Medication Assistance Plan DOCX if requested
            if generate_medication_plan:
                from create_final_tables import create_medication_assistance_plan_from_data
                from datetime import datetime
                import re
                
                first_name = pdf_data.get('First name (Details of the Client)', '').strip()
                surname = pdf_data.get('Surname (Details of the Client)', '').strip()
                dob_str = pdf_data.get('Date of birth (Details of the Client)', '').strip()
                ndis_number = pdf_data.get('NDIS number (Details of the Client)', '').strip()
                
                year = None
                if dob_str:
                    year_match = re.search(r'\b(19|20)\d{2}\b', dob_str)
                    if year_match:
                        year = year_match.group(0)
                if not year:
                    year = datetime.now().strftime('%Y')
                
                client_id = ''
                if ndis_number:
                    digits = re.sub(r'\D', '', ndis_number)
                    if len(digits) >= 6:
                        client_id = digits[:6]
                    elif len(digits) > 0:
                        client_id = digits.ljust(6, '0')
                if not client_id:
                    client_id = datetime.now().strftime('%H%M%S')
                
                name_part = f"{first_name} {surname}".strip() if (first_name or surname) else "test test"
                mp_filename = f"Medication Assistance Plan - {name_part} {year} - {client_id}.pdf"
                mp_path = os.path.join(app.config['UPLOAD_FOLDER'], mp_filename)
                create_medication_assistance_plan_from_data(pdf_data, mp_path, contact_name, active_users)
                output_files.append(('pdf', mp_path, mp_filename))
            
            # Clean up input file
            os.remove(filepath)
            
            # If only one output, send it directly
            if len(output_files) == 1:
                file_type, file_path, download_name = output_files[0]
                return send_file(file_path, as_attachment=True, download_name=download_name)
            else:
                # Multiple outputs - create a zip file
                zip_filename = f"outputs_{unique_filename}.zip"
                zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for file_type, file_path, download_name in output_files:
                        zipf.write(file_path, download_name)
                        os.remove(file_path)  # Clean up individual files
                response = send_file(zip_path, as_attachment=True, download_name='outputs.zip')
                # Clean up zip file after sending (Flask will handle cleanup, but we can also schedule it)
                return response
            
        except Exception as e:
            import traceback
            error_msg = str(e)
            error_traceback = traceback.format_exc()
            # Log full error for debugging (on Render, this goes to logs)
            print(f"ERROR processing file: {error_msg}")
            print(f"TRACEBACK: {error_traceback}")
            # Show user-friendly error message (limit length to avoid issues)
            if len(error_msg) > 200:
                error_msg = error_msg[:200] + "..."
            flash(f'Error processing file: {error_msg}')
            # Clean up files on error
            if 'filepath' in locals() and os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except:
                    pass
            return redirect(url_for('index'))
    else:
        flash('Invalid file type. Please upload a PDF file.')
        return redirect(url_for('index'))

# Verify critical imports at startup
def verify_imports():
    """Verify that required functions can be imported - helps catch deployment issues early"""
    try:
        import sys
        import traceback
        current_dir = os.getcwd()
        
        print("=" * 80)
        print("STARTUP: Verifying create_final_tables imports...")
        print(f"Current directory: {current_dir}")
        
        # Ensure paths are set up (same as in main code)
        if current_dir.endswith('/src') or current_dir.endswith('\\src'):
            parent_dir = os.path.dirname(current_dir)
            if parent_dir and parent_dir not in sys.path:
                sys.path.insert(0, parent_dir)
                print(f"Added parent directory to path: {parent_dir}")
        if current_dir not in sys.path:
            sys.path.insert(0, current_dir)
            print(f"Added current directory to path: {current_dir}")
        
        print(f"Python path (first 5): {sys.path[:5]}")
        
        # Check if file exists
        file_path = os.path.join(current_dir, 'create_final_tables.py')
        file_exists = os.path.exists(file_path)
        print(f"File exists at {file_path}: {file_exists}")
        
        if file_exists:
            file_size = os.path.getsize(file_path)
            print(f"File size: {file_size} bytes")
        
        # Try to import the module
        try:
            # Remove from cache if already loaded (might be broken)
            if 'create_final_tables' in sys.modules:
                del sys.modules['create_final_tables']
                print("Removed create_final_tables from cache")
            
            print("Attempting to import create_final_tables...")
            import create_final_tables
            module_file = getattr(create_final_tables, '__file__', 'unknown')
            print(f"✓ Module loaded from: {module_file}")
            
            # Check if functions exist
            has_parse = hasattr(create_final_tables, 'parse_pdf_to_data')
            has_ndis = hasattr(create_final_tables, 'load_ndis_support_items')
            has_users = hasattr(create_final_tables, 'load_active_users')
            
            print(f"Function checks:")
            print(f"  parse_pdf_to_data: {has_parse} (type: {type(getattr(create_final_tables, 'parse_pdf_to_data', None))})")
            print(f"  load_ndis_support_items: {has_ndis} (type: {type(getattr(create_final_tables, 'load_ndis_support_items', None))})")
            print(f"  load_active_users: {has_users} (type: {type(getattr(create_final_tables, 'load_active_users', None))})")
            
            if has_parse and has_ndis and has_users:
                print("✓ All required functions found in create_final_tables")
            else:
                print(f"⚠ WARNING: Some functions missing!")
                print(f"  Available attributes (first 20): {[a for a in dir(create_final_tables) if not a.startswith('_')][:20]}")
        except Exception as e:
            print(f"✗ ERROR: Could not import create_final_tables module")
            print(f"  Error type: {type(e).__name__}")
            print(f"  Error message: {str(e)}")
            print(f"  Full traceback:")
            print(traceback.format_exc())
            
            parent_dir = os.path.dirname(current_dir)
            if parent_dir:
                parent_file = os.path.join(parent_dir, 'create_final_tables.py')
                print(f"  File exists in parent ({parent_file}): {os.path.exists(parent_file)}")
        
        print("=" * 80)
    except Exception as e:
        import traceback
        print(f"✗ ERROR during import verification: {e}")
        print(f"  Traceback: {traceback.format_exc()}")

# Run verification when module loads (but don't fail if it doesn't work)
try:
    verify_imports()
except Exception:
    pass  # Don't prevent app from starting if verification fails

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)

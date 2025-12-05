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
    
    # Build display name from first name + surname
    display_name = " ".join([p for p in [first_name, surname] if p]).strip()
    
    # Use only home phone (work phone is not included)
    phone_number = home_phone if home_phone else ""

    return {
        "Salutation": "",
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
    
    if not generate_csv and not generate_service_agreement and not generate_emergency_plan and not generate_service_estimate and not generate_risk_assessment:
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
            from create_final_tables import parse_pdf_to_data
            pdf_data = parse_pdf_to_data(filepath)
            
            output_files = []
            
            # Generate CSV if requested
            if generate_csv:
                # Convert pdf_data format to the format expected by build_output_row
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
                    'email address': pdf_data.get('Email address (Contact Details of the Client)', '')
                }
                
                output_filename = f"transformed_{unique_filename}.csv"
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                with open(output_path, 'w', newline='', encoding='utf-8') as f_out:
                    writer = csv.DictWriter(f_out, fieldnames=OUTPUT_FIELDS)
                    writer.writeheader()
                    writer.writerow(build_output_row(parsed))
                output_files.append(('csv', output_path, 'client_export.csv'))
            
            # Generate Service Agreement PDF if requested
            if generate_service_agreement:
                # Import the service agreement generation function
                from create_final_tables import create_service_agreement_from_data
                # Get contact name from form
                contact_name = request.form.get('contact_name', '').strip()
                sa_filename = f"service_agreement_{unique_filename}.pdf"
                sa_path = os.path.join(app.config['UPLOAD_FOLDER'], sa_filename)
                # Pass source PDF path for signature extraction
                create_service_agreement_from_data(pdf_data, sa_path, contact_name, filepath)
                output_files.append(('pdf', sa_path, 'Service Agreement.pdf'))
            
            # Generate Emergency & Disaster Plan PDF if requested
            if generate_emergency_plan:
                # Import the emergency plan generation function
                from create_final_tables import create_emergency_disaster_plan_from_data
                # Get contact name from form (same as service agreement)
                contact_name = request.form.get('contact_name', '').strip()
                edp_filename = f"emergency_disaster_plan_{unique_filename}.pdf"
                edp_path = os.path.join(app.config['UPLOAD_FOLDER'], edp_filename)
                create_emergency_disaster_plan_from_data(pdf_data, edp_path, contact_name)
                output_files.append(('pdf', edp_path, 'Emergency & Disaster Plan.pdf'))
            
            # Generate Service Estimate Excel file if requested
            if generate_service_estimate:
                # Import the service estimate generation function
                from create_final_tables import create_service_estimate_csv
                # Get contact name from form (not used for Excel but kept for consistency)
                contact_name = request.form.get('contact_name', '').strip()
                se_filename = f"service_estimate_{unique_filename}.xlsx"
                se_path = os.path.join(app.config['UPLOAD_FOLDER'], se_filename)
                create_service_estimate_csv(pdf_data, se_path, contact_name)
                output_files.append(('xlsx', se_path, 'Service Estimate.xlsx'))
            
            # Generate Risk Assessment PDF if requested
            if generate_risk_assessment:
                # Import the risk assessment generation function
                from create_final_tables import create_risk_assessment_from_data
                # Get contact name from form for "Person Completing this assessment"
                contact_name = request.form.get('contact_name', '').strip()
                ra_filename = f"risk_assessment_{unique_filename}.pdf"
                ra_path = os.path.join(app.config['UPLOAD_FOLDER'], ra_filename)
                create_risk_assessment_from_data(pdf_data, ra_path, contact_name)
                output_files.append(('pdf', ra_path, 'Risk Assessment.pdf'))
            
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

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)

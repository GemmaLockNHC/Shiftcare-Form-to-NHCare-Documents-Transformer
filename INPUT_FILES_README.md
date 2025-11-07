# Required Input Files

This document describes the input files needed for the Service Agreement PDF generator.

## File Structure

All input files should be placed in the `outputs/other/` directory.

## Required Files

### 1. NDIS Support Items CSV
**File Path:** `outputs/other/NDIS Support Items - NDIS Support Items.csv`

**Required Columns:**
- `Support Item Name` - The name of the support item (used as lookup key)
- `Support Item Number` - The NDIS support item number
- `Unit` - The unit of measurement (e.g., "Hour", "Day")
- `WA` - The Western Australia price for the support item

**Example:**
```csv
Support Item Number,Support Item Name,Unit,WA
01_001_0107_1_1,Assistance With Self-Care Activities - Standard - Weekday Daytime,Hour,$65.47
01_001_0107_1_2,Assistance With Self-Care Activities - Standard - Weekday Evening,Hour,$70.00
```

**Note:** The CSV may have additional columns, but these four are required.

---

### 2. Active Users CSV
**File Path:** `outputs/other/Active_Users_1761707021.csv`

**Required Columns:**
- `name` - Full name of the user (used as lookup key)
- `mobile` - Mobile phone number
- `email` - Email address
- `area` or `role` - Team/area information (optional, used for 'team' field)

**Example:**
```csv
user_id,name,mobile,email,area,role
1,John Smith,0412345678,john.smith@example.com,Wanneroo,Coordinator
2,Jane Doe,0498765432,jane.doe@example.com,Perth,Manager
```

**Note:** The CSV may have additional columns, but `name`, `mobile`, and `email` are required.

---

### 3. Client Data Input (PDF or CSV)

The system can read client data from either a PDF or CSV file.

#### Option A: PDF Input (Preferred)
**File Path:** `outputs/other/Neighbourhood Care Welcoming Form Template 2.pdf`

This should be a PDF form with the following fields or text sections:

**Details of the Client Section:**
- First name
- Middle name
- Surname
- NDIS number
- Date of birth
- Gender

**Contact Details of the Client Section:**
- Home address
- Home phone
- Work phone
- Mobile phone
- Email address
- Preferred method of contact

**Additional Fields (used in PDF generation):**
- Total core budget to allocate to Neighbourhood Care
- Total capacity building budget to allocate to Neighbourhood Care
- Plan start date
- Plan end date
- Service start date
- Service end date
- Person signing the agreement
- First name (Person Signing the Agreement)
- Surname (Person Signing the Agreement)
- Relationship to client (Person Signing the Agreement)
- Home address (Person Signing the Agreement)
- First name (Primary carer)
- Surname (Primary carer)
- Relationship to client (Primary carer)
- Home address (Primary carer)
- First name (Emergency contact)
- Surname (Emergency contact)
- Is the primary carer also the emergency contact for the participant?
- Plan management type
- Plan manager name
- Plan manager postal address
- Plan manager phone number
- Plan manager email address
- Respondent
- Neighbourhood Care representative team
- Consent fields (various)

#### Option B: CSV Input (Fallback)
**File Path:** `outputs/other/Neighbourhood Care Welcoming Form Template 2.csv`  
**Alternative:** `outputs/other/Neighbourhood Care Welcoming Form.csv`

The CSV should have columns matching the field names above. The system will use the PDF parser first, then fall back to CSV if PDF parsing fails or doesn't yield enough data.

---

## Directory Structure

```
New JotForm/
├── outputs/
│   └── other/
│       ├── NDIS Support Items - NDIS Support Items.csv  (Required)
│       ├── Active_Users_1761707021.csv                   (Required)
│       ├── Neighbourhood Care Welcoming Form Template 2.pdf  (Optional - preferred input)
│       ├── Neighbourhood Care Welcoming Form Template 2.csv  (Optional - fallback)
│       └── Neighbourhood Care Welcoming Form.csv        (Optional - fallback)
├── app.py
├── create_final_tables.py
└── ...
```

---

## Notes

1. **NDIS Support Items CSV**: The system searches for support items by name. It first tries exact matching, then partial matching (case-insensitive). If not found, it returns placeholder values.

2. **Active Users CSV**: The system looks up users by the `name` field. It tries exact matching first, then partial matching. The `Respondent` field from the client data is used to look up the user.

3. **Client Data**: The system prioritizes PDF parsing. If the PDF has form fields, it extracts those. Otherwise, it extracts text and parses it section by section. CSV is only used as a fallback.

4. **File Names**: The exact file names matter! Make sure the files are named exactly as specified above.

---

## Testing

To verify your input files are set up correctly, you can run:

```bash
python3 create_final_tables.py
```

This will:
- Load the NDIS support items
- Load the active users
- Parse the client data (PDF or CSV)
- Generate the Service Agreement PDF

If any required files are missing, you'll see error messages indicating which files need to be added.


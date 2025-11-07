# Deployment Checklist for Web App

## Files That MUST Be in GitHub

### Core Application Files
- ✅ `app.py` - Main Flask application
- ✅ `create_final_tables.py` - PDF generation logic
- ✅ `templates/index.html` - Web interface
- ✅ `requirements.txt` - Python dependencies
- ✅ `runtime.txt` - Python version (if using Heroku/Railway)

### Required Input Data Files (Now Allowed by .gitignore)
- ✅ `outputs/other/NDIS Support Items - NDIS Support Items.csv` - **REQUIRED**
- ✅ `outputs/other/Active_Users_1761707021.csv` - **REQUIRED**
- ⚠️ `outputs/other/Neighbourhood Care Welcoming Form Template 2.pdf` - Optional (users upload their own)
- ⚠️ `outputs/other/Neighbourhood Care Welcoming Form Template 2.csv` - Optional (fallback)

### Documentation
- ✅ `README.md`
- ✅ `INPUT_FILES_README.md`
- ✅ `GITHUB_SETUP.md`
- ✅ `FILE_STRUCTURE.md`

## Files That Should NOT Be in GitHub

- ❌ `uploads/` - User-uploaded files (created at runtime)
- ❌ `__pycache__/` - Python cache
- ❌ `*.pyc` - Compiled Python files
- ❌ `Service Agreement - FINAL TABLES.pdf` - Generated output files

## Quick Fix Steps

1. **Update .gitignore** (✅ Already done!)
   - The `.gitignore` now allows the required input files

2. **Add the required files to git:**
   ```bash
   git add outputs/other/NDIS\ Support\ Items\ -\ NDIS\ Support\ Items.csv
   git add outputs/other/Active_Users_1761707021.csv
   git add .gitignore
   git add app.py create_final_tables.py templates/index.html requirements.txt
   git commit -m "Add required input files and update gitignore"
   git push
   ```

3. **Verify files are tracked:**
   ```bash
   git ls-files | grep -E "(NDIS|Active_Users)"
   ```

## What the Web App Needs to Run

The web app (`app.py`) needs these files to exist on the server:

1. **`outputs/other/NDIS Support Items - NDIS Support Items.csv`**
   - Used by `create_final_tables.py` to lookup support item prices
   - **MUST be present** or PDF generation will fail

2. **`outputs/other/Active_Users_1761707021.csv`**
   - Used to lookup Neighbourhood Care staff contact details
   - **MUST be present** or will show placeholder data

3. **User-uploaded PDFs** (handled automatically)
   - Users upload PDFs through the web interface
   - Stored temporarily in `uploads/` folder

## Testing Locally

Before deploying, test that files are accessible:

```bash
python3 -c "
import os
required = [
    'outputs/other/NDIS Support Items - NDIS Support Items.csv',
    'outputs/other/Active_Users_1761707021.csv'
]
for f in required:
    if os.path.exists(f):
        print(f'✅ {f}')
    else:
        print(f'❌ {f} MISSING')
"
```

## Deployment Platforms

### Heroku
- Files in GitHub are automatically deployed
- Make sure `outputs/other/` directory structure is preserved
- May need to create directory structure if it doesn't exist

### Railway / Render / Other
- Same as Heroku - files from GitHub are deployed
- Ensure the directory structure `outputs/other/` exists

## If Files Are Still Missing After Deployment

1. **Check the deployed file structure:**
   - SSH into your server or check logs
   - Verify `outputs/other/` directory exists

2. **Manually upload files:**
   - Use your platform's file upload feature
   - Or use `git` to push the files directly

3. **Check file paths:**
   - The code uses relative paths: `outputs/other/...`
   - Make sure the working directory is correct when the app runs


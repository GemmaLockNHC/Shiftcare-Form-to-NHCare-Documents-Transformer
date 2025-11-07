# How to Connect This Folder to a GitHub Repository

## Step 1: Initialize Git in This Folder

Open Terminal and run:

```bash
cd ~/Desktop/"JotForm Service Agreement"
git init
```

## Step 2: Add All Files

```bash
git add .
```

## Step 3: Make Your First Commit

```bash
git commit -m "Initial commit - Service Agreement Generator"
```

## Step 4: Create a New Repository on GitHub

1. Go to [GitHub.com](https://github.com)
2. Click the **+** icon in the top right
3. Select **"New repository"**
4. Name it (e.g., `jotform-service-agreement`)
5. **DO NOT** initialize with README, .gitignore, or license (we already have these)
6. Click **"Create repository"**

## Step 5: Connect Your Local Folder to GitHub

GitHub will show you commands. Use these:

```bash
# Add the remote repository (replace YOUR_USERNAME and REPO_NAME)
git remote add origin https://github.com/YOUR_USERNAME/REPO_NAME.git

# Rename the default branch to main (if needed)
git branch -M main

# Push your code to GitHub
git push -u origin main
```

## Step 6: Verify

Go to your GitHub repository page and you should see all your files!

## Future Updates

When you make changes:

```bash
git add .
git commit -m "Description of your changes"
git push
```

## Troubleshooting

### If you get "remote origin already exists"
```bash
git remote remove origin
git remote add origin https://github.com/YOUR_USERNAME/REPO_NAME.git
```

### If you need to force push (be careful!)
```bash
git push -u origin main --force
```

### Check what files will be committed
```bash
git status
```

## Files That Will Be Uploaded

✅ All code files (app.py, create_final_tables.py, etc.)
✅ Templates (index.html)
✅ Required input files (NDIS Support Items CSV, Active Users CSV)
✅ Configuration files (requirements.txt, runtime.txt, .gitignore)
✅ Documentation files

❌ User uploads (uploads/ folder is ignored)
❌ Generated PDFs (ignored by .gitignore)


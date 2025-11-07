# How to Push to GitHub

Your repository is set up and ready! You just need to authenticate to push.

## Option 1: Use GitHub Desktop (Easiest)

1. Download GitHub Desktop: https://desktop.github.com/
2. Sign in with your GitHub account
3. File → Add Local Repository
4. Choose: `~/Desktop/JotForm Service Agreement`
5. Click "Publish repository" button

## Option 2: Use Personal Access Token (Terminal)

### Step 1: Create a Personal Access Token

1. Go to GitHub.com → Settings → Developer settings → Personal access tokens → Tokens (classic)
2. Click "Generate new token (classic)"
3. Give it a name like "JotForm Service Agreement"
4. Select scopes: Check **"repo"** (this gives full repository access)
5. Click "Generate token"
6. **COPY THE TOKEN** (you won't see it again!)

### Step 2: Push Using the Token

Run this command (it will ask for your username and password):

```bash
cd ~/Desktop/"JotForm Service Agreement"
git push -u origin main
```

When it asks for:
- **Username**: `GemmaLockNHC`
- **Password**: Paste your Personal Access Token (not your GitHub password!)

## Option 3: Use GitHub CLI (if installed)

```bash
gh auth login
cd ~/Desktop/"JotForm Service Agreement"
git push -u origin main
```

## Quick Check

To verify everything is ready:

```bash
cd ~/Desktop/"JotForm Service Agreement"
git status
git remote -v
```

You should see:
- `origin https://github.com/GemmaLockNHC/Form.git`
- All files committed

Then just push when you're authenticated!


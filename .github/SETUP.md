# GitHub Actions Setup Guide

## Required Repository Secrets

To enable the Windows build workflow to push to your private repository, you need to configure the following secrets in your GitHub repository settings:

### 1. PRIVATE_REPO_TOKEN
- **Purpose**: Personal Access Token for authenticating with the private repository
- **How to create**:
  1. Go to GitHub Settings → Developer settings → Personal access tokens → Tokens (classic)
  2. Click "Generate new token (classic)"
  3. Select scopes: `repo` (Full control of private repositories)
  4. Copy the generated token

### Repository Configuration
- **Private Repository**: `makromaster/rnb-snow-builds`
- **Note**: The repository URL is now hardcoded in the workflow, so you only need the PRIVATE_REPO_TOKEN secret

## Setting Up Repository Secrets

1. Go to your repository on GitHub
2. Click on "Settings" tab
3. In the left sidebar, click "Secrets and variables" → "Actions"
4. Click "New repository secret"
5. Add the secret:
   - Name: `PRIVATE_REPO_TOKEN`
   - Value: Your GitHub Personal Access Token (configured automatically)

## Workflow Behavior

- **Triggers**: Runs on pushes to `main`/`develop` branches and pull requests to `main`
- **Windows Build**: Tests the application on Windows environment
- **Artifact Upload**: Creates downloadable build artifacts
- **Private Repo Push**: Only pushes to private repo on `main` branch pushes (not PRs)

## Security Notes

- Never commit tokens or sensitive URLs directly to your repository
- The workflow only pushes to the private repo on main branch commits
- All secrets are encrypted and only accessible to GitHub Actions

## Customization

You can modify the workflow file (`.github/workflows/windows-build.yml`) to:
- Change Python version
- Add additional test steps
- Modify the build artifact contents
- Change the target branch for private repo pushes
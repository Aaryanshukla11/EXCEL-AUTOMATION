# Git Commit & Remote Automation Tools

This directory contains dedicated automation scripts for managing Git commit histories, backdating commits across project development phases (1st Nov – 20th Nov 2026), and pushing branches to remote repositories (GitHub/GitLab).

## Files in this Directory

- `rebuild_dated_commits.py`: Re-dates and aligns commit history progressively from Nov 1 to Nov 20, 2026 across all project phases.
- `push_to_remote.py`: Interactive script to set the remote repository URL and push both `main` and `master` branches with tags/commits intact.

## How to Push to GitHub

Run the following command in terminal:
```bash
python gitcommit/push_to_remote.py --url <YOUR_GITHUB_REPO_URL>
```
Or interactively:
```bash
python gitcommit/push_to_remote.py
```

# ScriptRunner Scripts Repository

This repository contains Python scripts managed by [ScriptRunner](https://github.com/PatrykEmilLibert/script-runner).

## Structure

```
scripts/
├── script_name/
│   ├── main.py              # Main script file
│   ├── requirements.txt     # Python dependencies
│   └── metadata.json        # Script metadata
```

## Features

- **Automatic dependency tracking** - Requirements are auto-detected from imports
- **Version control** - All scripts are backed up to GitHub
- **Team collaboration** - Share scripts with your team

## Usage

Scripts are automatically synced from this repository to ScriptRunner. Add new scripts using the "Add Script" button in the application.

## Script Metadata

Each script includes a `metadata.json` file:

```json
{
  "name": "Script Name",
  "description": "What the script does",
  "author": "Your Name",
  "version": "1.0.0",
  "created_at": "2026-01-20T00:00:00Z",
  "last_modified": "2026-01-20T00:00:00Z"
}
```

## Adding Scripts

Use ScriptRunner application to add scripts. The app will:
1. Analyze your Python code
2. Auto-detect dependencies
3. Generate `requirements.txt`
4. Commit and push to this repository

---

Managed by ScriptRunner v0.1.1

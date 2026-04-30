# MoamProject — Mail Generator

Internal tool for generating structured product mails at Van Lanschot Kempen.

## Features
- **Documentatie Mail** — documentation confirmation emails
- **PC Mail** — product notification emails with Word + Excel attachments
- **Marketing Mail** — marketing distribution emails
- **Increase / Decrease Mail** — size adjustment notifications
- **SharePoint integration** — load live deal data directly from SharePoint summary

## Setup

```bash
# Create virtual environment
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Running from source

```bash
python -m gui.app
```

## Building the EXE

```powershell
powershell -ExecutionPolicy Bypass -File build_exe.ps1
```

Output: `dist\MoamProject.exe`

## Deploying to shared disk

Copy `dist\MoamProject.exe` to the shared drive.  
Users run the EXE directly — on first run it extracts to local temp, subsequent launches are fast.

## SharePoint data

The app reads deal data from `sharepoint\sharepointsummary.xlsx`.  
Use the **Refresh SharePoint** button in the toolbar to pull the latest data.

# build_exe.ps1
$ErrorActionPreference = "Stop"

# ----- CONFIG (edit these 3 lines) -----
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython  = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$EntryScript = Join-Path $ProjectRoot "gui\app.py"   # <-- your entrypoint

# Optional: app identity
$AppName     = "MoamProject"
$IconPath    = Join-Path $ProjectRoot "assets\app.ico"  # optional; ok if missing

# ----- SANITY CHECKS -----
if (!(Test-Path $VenvPython)) { throw "Python venv not found at: $VenvPython" }
if (!(Test-Path $EntryScript)) { throw "Entry script not found at: $EntryScript" }

Write-Host "Using Python: $VenvPython"
Write-Host "Entry script: $EntryScript"
Write-Host "Project root: $ProjectRoot"

# ----- INSTALL/UPDATE PYINSTALLER -----
& $VenvPython -m pip install --upgrade pip | Out-Host
& $VenvPython -m pip install --upgrade pyinstaller | Out-Host
& $VenvPython -m pip install --upgrade pillow | Out-Host


# ----- CLEAN OLD BUILDS -----
$buildDir = Join-Path $ProjectRoot "build"
$distDir  = Join-Path $ProjectRoot "dist"
$specFile = Join-Path $ProjectRoot "$AppName.spec"

if (Test-Path $buildDir) { Remove-Item $buildDir -Recurse -Force }
if (Test-Path $distDir)  { Remove-Item $distDir  -Recurse -Force }
if (Test-Path $specFile) { Remove-Item $specFile -Force }

# ----- PYINSTALLER ARGS -----
$Args = @(
  "--noconfirm",
  "--clean",
  "--onefile",
  "--windowed",              # no console window (Tkinter GUI)
  "--name", $AppName,
  "--paths", $ProjectRoot,

  # --- DATA FILES ---
  "--add-data", "$(Join-Path $ProjectRoot 'statics\static_sheet.xlsx');statics",
  "--add-data", "$(Join-Path $ProjectRoot 'sharepoint\sharepointsummary.xlsx');sharepoint",
  "--add-data", "$(Join-Path $ProjectRoot 'PCMail\template\Individual Product Notification - template.docx');PCMail\template",
  "--add-data", "$(Join-Path $ProjectRoot 'PCMail\template\pcmailtargetmarkets.xlsx');PCMail\template",
  "--add-data", "$(Join-Path $ProjectRoot 'MarketingMail\templates\excel(marketingmail).xlsx');MarketingMail\templates",
  "--add-data", "$(Join-Path $ProjectRoot 'increase_decrease_mail\increasedecrease(mailtemplate).xlsx');increase_decrease_mail",
  "--add-data", "$(Join-Path $ProjectRoot 'assets\app.ico');assets",

  $EntryScript
)

# Add icon if present
if (Test-Path $IconPath) {
  $Args = @("--icon", $IconPath) + $Args
  Write-Host "Icon: $IconPath"
} else {
  Write-Host "No icon found at $IconPath (skipping icon)."
}

# ----- BUILD -----
Write-Host "Building EXE..."
& $VenvPython -m PyInstaller @Args | Out-Host

$exePath = Join-Path $distDir "$AppName.exe"
if (!(Test-Path $exePath)) { throw "Build finished but EXE not found: $exePath" }

Write-Host ""
Write-Host "Done!"
Write-Host "EXE: $exePath"
Write-Host ""
Write-Host "Deploy: copy $AppName.exe to the shared disk."
Write-Host "On first run it extracts to local temp (once per version), then runs fast."
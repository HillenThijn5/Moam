# build_exe.ps1
$ErrorActionPreference = "Stop"

# ----- CONFIG (edit these 3 lines) -----
$ProjectRoot = $PSScriptRoot
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
$PyArgs = @(
  "--noconfirm",
  "--clean",
  "--onefile",
  "--windowed",              # no console window (Tkinter GUI)
  "--name", $AppName,
  "--paths", $ProjectRoot,
  "--distpath", $distDir,    # force output into project's dist\
  "--workpath", $buildDir,   # force build cache into project's build\
  "--specpath", $ProjectRoot,# put the .spec file in the project root

  # --- BUNDLED RESOURCES (icon only — xlsx/docx live beside the EXE) ---
  "--add-data", "$(Join-Path $ProjectRoot 'assets\app.ico');assets",

  $EntryScript
)

# Add icon if present
if (Test-Path $IconPath) {
  $PyArgs = @("--icon", $IconPath) + $PyArgs
  Write-Host "Icon: $IconPath"
} else {
  Write-Host "No icon found at $IconPath (skipping icon)."
}

# ----- BUILD -----
Write-Host "Building EXE..."
& $VenvPython -m PyInstaller @PyArgs | Out-Host

$exePath = Join-Path $distDir "$AppName.exe"
if (!(Test-Path $exePath)) { throw "Build finished but EXE not found: $exePath" }


# ----- COPY EXTERNAL DATA FILES BESIDE THE EXE -----
Write-Host "Copying external templates and Excel sheets beside the EXE..."

$externalFiles = @(
  @{ Src = "statics\static_sheet.xlsx";                                          Dst = "statics\static_sheet.xlsx" },
  @{ Src = "sharepoint\sharepointsummary.xlsx";                                  Dst = "sharepoint\sharepointsummary.xlsx" },
  @{ Src = "PCMail\template\Individual Product Notification - template.docx";    Dst = "PCMail\template\Individual Product Notification - template.docx" },
  @{ Src = "PCMail\template\pcmailtargetmarkets.xlsx";                           Dst = "PCMail\template\pcmailtargetmarkets.xlsx" },
  @{ Src = "MarketingMail\templates\excel(marketingmail).xlsx";                  Dst = "MarketingMail\templates\excel(marketingmail).xlsx" },
  @{ Src = "increase_decrease_mail\increasedecrease(mailtemplate).xlsx";         Dst = "increase_decrease_mail\increasedecrease(mailtemplate).xlsx" }
)

foreach ($f in $externalFiles) {
  $src = Join-Path $ProjectRoot $f.Src
  $dst = Join-Path $distDir     $f.Dst
  if (!(Test-Path $src)) { Write-Warning "Source not found, skipping: $src"; continue }
  $dstDir = Split-Path $dst -Parent
  if (!(Test-Path $dstDir)) { New-Item $dstDir -ItemType Directory -Force | Out-Null }
  Copy-Item $src $dst -Force
  Write-Host "  Copied: $($f.Dst)"
}

Write-Host ""
Write-Host "Done!"
Write-Host "EXE: $exePath"
Write-Host ""
Write-Host "Deploy: copy the entire 'dist' folder to the shared disk."
Write-Host "The EXE and all template/Excel files must stay together."
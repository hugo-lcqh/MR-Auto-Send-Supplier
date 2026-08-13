$root = Split-Path -Parent $PSScriptRoot
$launcherPath = Join-Path $root "Scripts\MR-Launcher.ps1"
$installerPath = Join-Path $root "Installer\MR-Auto-Setup.iss"
$releaseNotesPath = Join-Path $root "Installer\ReleaseNotes.txt"
$sampleSupplierPath = Join-Path $root "Config\suppliers.example.csv"

$tokens = $null
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($launcherPath, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw "MR-Launcher.ps1 has parser errors" }

$launcher = Get-Content -Raw -LiteralPath $launcherPath
$launcherRequirements = @(
    '$appVersion = "1.3.0"',
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{8D183A16-6C3F-4EF1-B296-502FB04CF3C6}_is1',
    'DisplayVersion',
    'InstallLocation',
    '[IO.Path]::GetFullPath',
    'Version $appVersion | Developed by Hugo Le Chi Quoc Hung'
)
$missingLauncher = @($launcherRequirements | Where-Object { !$launcher.Contains($_) })
if ($missingLauncher.Count) {
    throw "Launcher version/developer information missing: $($missingLauncher -join ', ')"
}

$installer = Get-Content -Raw -LiteralPath $installerPath
$installerRequirements = @(
    '#define MyAppVersion "1.3.0"',
    '#define MyAppPublisher "Hugo Aka Le Chi Quoc Hung"',
    'OutputBaseFilename=MR-Auto-Setup-v{#MyAppVersion}',
    'VersionInfoVersion={#MyAppVersion}.0',
    'VersionInfoCompany={#MyAppPublisher}',
    'InfoBeforeFile=ReleaseNotes.txt',
    'Source: "..\Config\suppliers.example.csv"; DestDir: "{app}\Config"; DestName: "suppliers.csv"',
    'WizardInfoBefore=What''s New in MR Auto Send Supplier',
    'InfoBeforeLabel=Please review the fixes and improvements in version {#MyAppVersion} before continuing.'
)
$missingInstaller = @($installerRequirements | Where-Object { !$installer.Contains($_) })
if ($missingInstaller.Count) {
    throw "Installer version/developer metadata missing: $($missingInstaller -join ', ')"
}

if (!(Test-Path -LiteralPath $releaseNotesPath)) {
    throw "Installer release notes file is missing: $releaseNotesPath"
}
$releaseNotes = [Text.Encoding]::UTF8.GetString([IO.File]::ReadAllBytes($releaseNotesPath))
$releaseNotesRequirements = @(
    'VERSION 1.3.0',
    '13-Aug-2026',
    'Material Request Control Center',
    'Supplier Queue',
    'Operation Workflow',
    'Display only, do not send',
    'suppliers.csv',
    'reply_folder.txt'
)
$missingReleaseNotes = @($releaseNotesRequirements | Where-Object { !$releaseNotes.Contains($_) })
if ($missingReleaseNotes.Count) {
    throw "Installer release notes are incomplete: $($missingReleaseNotes -join ', ')"
}

if (!(Test-Path -LiteralPath $sampleSupplierPath)) {
    throw "Safe sample supplier configuration is missing: $sampleSupplierPath"
}
$sampleSuppliers = @(Import-Csv -LiteralPath $sampleSupplierPath)
$requiredSupplierColumns = @("Keyword", "VendorCode", "VendorName", "Email To", "Email CC", "MC", "Buyer")
$sampleSupplierColumns = @($sampleSuppliers[0].PSObject.Properties.Name | ForEach-Object { $_.Trim() })
$missingSupplierColumns = @($requiredSupplierColumns | Where-Object { $sampleSupplierColumns -notcontains $_ })
if (!$sampleSuppliers.Count -or $missingSupplierColumns.Count) {
    throw "Sample supplier configuration is invalid; missing columns: $($missingSupplierColumns -join ', ')"
}
foreach ($sampleSupplier in $sampleSuppliers) {
    foreach ($emailColumn in @("Email To", "Email CC", "MC ")) {
        foreach ($address in ([string]$sampleSupplier.$emailColumn -split ';')) {
            $address = $address.Trim()
            if ($address -and $address -notmatch '@example\.(com|net|org)$') {
                throw "Sample supplier configuration must contain reserved example addresses only"
            }
        }
    }
}

Write-Host "PASS: Launcher, Setup, and release notes show synchronized version information"

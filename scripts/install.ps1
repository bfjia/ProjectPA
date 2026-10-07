# Installs the Release build for the current user. No admin. Never touches the Outlook process:
# the new build loads the next time Outlook starts.
param([switch]$SkipChecks)   # install even when a dependency is missing
$ErrorActionPreference = 'Stop'
$src = Join-Path (Split-Path $PSScriptRoot) 'src\ProjectPA.AddIn\bin\Release\net48'

# ---- dependencies: collect every problem, explain each, change nothing unless all pass ----
$missing = @()
function Need($what, $how) { $script:missing += "- $what`n    $how" }
function Reg($key, $name) { try { (Get-ItemProperty $key -ErrorAction Stop).$name } catch { $null } }

if (-not [Environment]::Is64BitProcess) {
    Need 'This is 32-bit PowerShell.' 'Run the script from 64-bit Windows PowerShell, so the registry keys land where 64-bit Outlook looks.'
}

$outlook = Reg 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' '(default)'
if (-not $outlook -or -not (Test-Path $outlook)) {
    Need 'Classic Outlook is not installed.' 'Install Outlook from Microsoft 365 (or Office 2016 or later), 64-bit. The "new Outlook" app cannot load this add-in.'
} elseif ([Environment]::Is64BitProcess -and (Reg 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook' 'Bitness') -ne 'x64') {
    # only trustworthy from 64-bit PowerShell: a 32-bit one is shown a different part of the registry
    Need 'Outlook is not the 64-bit edition.' 'This script registers for 64-bit Outlook only. Reinstall Office as 64-bit (account.microsoft.com, Install, Other options).'
}

if ((Reg 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' 'Release') -lt 528040) {
    Need '.NET Framework 4.8 is missing.' 'Run Windows Update, or install it from https://dotnet.microsoft.com/download/dotnet-framework/net48'
}

if (-not (Test-Path "$src\ProjectPA.AddIn.dll")) {
    Need 'The add-in has not been built.' 'Run scripts\build.ps1 first.'
    # no build yet, so say now what the build will need
    if (-not (Get-Command dotnet -ErrorAction SilentlyContinue) -or -not (dotnet --list-sdks)) {
        Need 'The .NET SDK is missing (needed to build).' 'Install .NET SDK 9 from https://dotnet.microsoft.com/download'
    }
    if (-not (Test-Path 'C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook')) {
        Need 'The Office interop assemblies are missing (needed to build).' 'In Office setup (Apps, Microsoft 365, Modify) make sure ".NET programmability support" is included, then repair Office.'
    }
}

# Claude Code, found the way the add-in finds it
$saved = try { (Get-Content "$env:LOCALAPPDATA\ProjectPA\settings.json" -Raw | ConvertFrom-Json).ClaudePath } catch { $null }
$claude = @(
    $saved
    (Get-Command claude.exe -ErrorAction SilentlyContinue).Source
    "$env:USERPROFILE\.local\bin\claude.exe"
    (Get-ChildItem "$env:USERPROFILE\.vscode\extensions\anthropic.claude-code-*\resources\native-binary\claude.exe" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
if (-not $claude) {
    Need 'Claude Code is not installed.' 'Install it (https://code.claude.com/docs/en/quickstart) or the Claude Code extension for VS Code, then run "claude" once and sign in with your Claude subscription.'
} else {
    $auth = try { & $claude auth status | Out-String | ConvertFrom-Json } catch { $null }
    if (-not $auth.loggedIn) {
        Need 'Claude Code is not signed in.' "Run `"$claude`" in a terminal and sign in with your Claude account (a Pro or Max subscription)."
    }
}

if ($missing -and -not $SkipChecks) {
    "Cannot install yet. $($missing.Count) thing(s) to fix:`n"
    $missing
    "`nNothing was changed. Fix the above and run this script again (or pass -SkipChecks to install anyway)."
    exit 1
}
"Checks passed: 64-bit Outlook, .NET Framework 4.8, build present, Claude Code signed in ($($auth.subscriptionType))."

# each build gets its own folder, so copying works while Outlook holds the previous one
$app = Join-Path $env:LOCALAPPDATA 'ProjectPA\app'
$dst = Join-Path $app (Get-Date -Format 'yyyyMMdd-HHmmss')
New-Item -ItemType Directory -Force $dst | Out-Null
Copy-Item "$src\*" $dst -Recurse

$dll = Join-Path $dst 'ProjectPA.AddIn.dll'
$asm = [Reflection.AssemblyName]::GetAssemblyName($dll).FullName
$classes = @{
    'ProjectPA.Connect'  = '{7C8FF91B-71C4-40E7-AF0F-813F89238FC8}', 'ProjectPA.AddIn.Connect'
    'ProjectPA.PaneHost' = '{84A220F1-8ABF-4080-9B87-ECE5C86C5F67}', 'ProjectPA.AddIn.PaneHost'
}
foreach ($progId in $classes.Keys) {
    $guid, $class = $classes[$progId]
    $p = "HKCU:\Software\Classes\$progId"
    $c = "HKCU:\Software\Classes\CLSID\$guid"
    $s = "$c\InprocServer32"
    New-Item "$p\CLSID", "$c\ProgId", $s -Force | Out-Null
    Set-ItemProperty $p '(default)' $class
    Set-ItemProperty "$p\CLSID" '(default)' $guid
    Set-ItemProperty $c '(default)' $class
    Set-ItemProperty "$c\ProgId" '(default)' $progId
    # same values regasm /codebase writes, under HKCU
    Set-ItemProperty $s '(default)' 'mscoree.dll'
    Set-ItemProperty $s ThreadingModel 'Both'
    Set-ItemProperty $s Class $class
    Set-ItemProperty $s Assembly $asm
    Set-ItemProperty $s RuntimeVersion 'v4.0.30319'
    Set-ItemProperty $s CodeBase ([Uri]$dll).AbsoluteUri
}

$addin = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\ProjectPA.Connect'
New-Item $addin -Force | Out-Null
Set-ItemProperty $addin FriendlyName 'PApii'
Set-ItemProperty $addin Description 'PApii, a personal assistant for email, powered by Claude Code.'
Set-ItemProperty $addin LoadBehavior 3 -Type DWord

# old builds: only when Outlook is closed, it may still be loading files from them
if (-not (Get-Process OUTLOOK -ErrorAction SilentlyContinue)) {
    Get-ChildItem $app -Directory | Where-Object FullName -ne $dst | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

"Installed $dst"
'Restart Outlook to load it.'

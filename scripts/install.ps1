# Installs the Release build for the current user. No admin. Never touches the Outlook process:
# the new build loads the next time Outlook starts.
$ErrorActionPreference = 'Stop'
$src = Join-Path (Split-Path $PSScriptRoot) 'src\ProjectPA.AddIn\bin\Release\net48'
if (-not (Test-Path "$src\ProjectPA.AddIn.dll")) { throw 'Nothing to install. Run scripts\build.ps1 first.' }
if (-not [Environment]::Is64BitProcess) { throw 'Run from 64-bit PowerShell so the keys land where 64-bit Outlook looks.' }

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

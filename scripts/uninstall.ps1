# Removes the add-in registration. Settings, logs and sessions stay unless -PurgeData is given.
param([switch]$PurgeData)

$keys = 'HKCU:\Software\Microsoft\Office\Outlook\Addins\ProjectPA.Connect',
        'HKCU:\Software\Classes\ProjectPA.Connect',
        'HKCU:\Software\Classes\ProjectPA.PaneHost',
        'HKCU:\Software\Classes\CLSID\{7C8FF91B-71C4-40E7-AF0F-813F89238FC8}',
        'HKCU:\Software\Classes\CLSID\{84A220F1-8ABF-4080-9B87-ECE5C86C5F67}'
$keys | Where-Object { Test-Path $_ } | Remove-Item -Recurse -Force

$data = Join-Path $env:LOCALAPPDATA 'ProjectPA'
if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
    'Registration removed. Outlook is running: close it, then run this again to delete the files.'
} else {
    if ($PurgeData) {
        # Claude Code keeps its own transcript of each saved session: have it drop those before the folders go
        $saved = try { (Get-Content "$data\settings.json" -Raw | ConvertFrom-Json).ClaudePath } catch { $null }
        $claude = @(
            $saved
            (Get-Command claude.exe -ErrorAction SilentlyContinue).Source
            "$env:USERPROFILE\.local\bin\claude.exe"
            (Get-ChildItem "$env:USERPROFILE\.vscode\extensions\anthropic.claude-code-*\resources\native-binary\claude.exe" -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
        ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
        if ($claude) {
            Get-ChildItem "$data\sessions" -Directory -ErrorAction SilentlyContinue | ForEach-Object { & $claude purge $_.FullName --yes | Out-Null }
        }
    }
    $target = if ($PurgeData) { $data } else { Join-Path $data 'app' }
    if (Test-Path $target) { Remove-Item $target -Recurse -Force }
    "Removed $target"
}

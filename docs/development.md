# Development

This document covers building, testing, installing and debugging ProjectPA. The add-in it builds is called PApii in Outlook; code, folders and registry names use ProjectPA.

## Prerequisites

Everything below was already present on the development machine; nothing is installed by the build.

- Windows with classic Outlook (Microsoft 365, 64-bit).
- .NET SDK 9 and the .NET Framework 4.8 targeting pack.
- The Office interop assemblies in the Global Assembly Cache (`C:\Windows\assembly`). Office installs them. The add-in project references them by path and embeds the types it uses, so no interop DLL is shipped.
- Claude Code, signed in with a Claude subscription.
- Git. On the development machine git exists only inside WSL (see "Git" below).

NuGet packages restored from nuget.org: `Newtonsoft.Json` for the product, and `xunit`, `xunit.runner.visualstudio` and `Microsoft.NET.Test.Sdk` for the tests.

## Layout

| Path | Purpose |
|---|---|
| `src/ProjectPA.Core` | Logic with no Outlook or UI dependency: the Claude runner, the thread model and its rendering, attachment text extraction, prompts, settings, logging. |
| `src/ProjectPA.UI` | The WPF pane, the settings and prompts windows, and the `PApii` class that drives the pane. No Outlook dependency. |
| `src/ProjectPA.AddIn` | The COM add-in Outlook loads: ribbon, task pane hosting, and all Outlook object-model code. |
| `src/ProjectPA.DevHost` | A small executable that shows the same pane in a normal window, for working without Outlook. |
| `tests/ProjectPA.Core.Tests` | Unit tests for the core. |
| `scripts` | Build, install and uninstall scripts. |

All projects target .NET Framework 4.8, because that is the runtime classic Outlook loads add-ins into. Shared build settings are in `Directory.Build.props`.

## Build and test

```powershell
scripts\build.ps1
```

This runs `dotnet build ProjectPA.sln -c Release` followed by `dotnet test`. Both must pass before installing.

## Install into Outlook

```powershell
scripts\install.ps1
```

The script needs no administrator rights and does three things:

1. Copies the Release build to a new folder, `%LOCALAPPDATA%\ProjectPA\app\<timestamp>`. Each build gets its own folder so that installing works while Outlook still has the previous build loaded.
2. Registers the two COM classes under `HKCU\Software\Classes`: `ProjectPA.Connect` (the add-in) and `ProjectPA.PaneHost` (the control shown in the task pane). The values are the same ones `regasm /codebase` would write, but per user.
3. Registers the add-in with Outlook under `HKCU\Software\Microsoft\Office\Outlook\Addins\ProjectPA.Connect` with `LoadBehavior` 3 (load at startup).

The script never closes or restarts Outlook. The new build is picked up the next time Outlook starts, so **restart Outlook yourself after installing**. Older build folders are deleted on a later install, and only when Outlook is not running.

To remove the add-in:

```powershell
scripts\uninstall.ps1              # removes registration and the installed files
scripts\uninstall.ps1 -PurgeData   # also deletes settings, logs and saved sessions
```

### Checking the registration without Outlook

Both classes can be created from 64-bit PowerShell, which proves the registry entries and the build are sound before Outlook is involved:

```powershell
$c = New-Object -ComObject ProjectPA.Connect
$c.GetCustomUI('Microsoft.Outlook.Explorer')   # prints the ribbon XML
New-Object -ComObject ProjectPA.PaneHost       # constructs the pane
```

## Working without Outlook: the DevHost

`ProjectPA.DevHost.exe` hosts the real pane and makes real Claude calls. It is the fastest way to work on the pane, the prompts and the Claude integration.

By default it works on a built-in sample thread (a three-message budget discussion).

| Option | Effect |
|---|---|
| `--do assist\|draft\|summarize\|times\|event\|tasks\|brief\|polish\|follow` | Run that action on start. `times` is Find Times, `event` is Add to Calendar, `brief` is Write from Brief, `follow` is Follow Up in 3 days. Calendar and task actions use a made-up calendar, and what they "create" is written to `out.png.events.txt` when `--shot` is used. |
| `--with "text"` | Instructions for `--do draft`, or the brief for `--do brief`. |
| `--typed "text"` | Pretend this text is being written in a reply, for `polish` and `brief`. |
| `--selected "text"` | Pretend this part of it is highlighted, for `polish`. |
| `--click "Label"` | Then press the button or pill with that label on the last card, for example `Shorter` or `Reply All`. |
| `--say "text"` | Then type the text into the pane and send it. |
| `--attach file` | Add a file to the sample thread as an attachment. |
| `--show prompts\|settings` | Show the prompts window or the settings window instead of the pane. |
| `--model haiku` | Model for this run only; not saved. |
| `--shot out.png` | When everything has finished, save a picture of the pane (or of the `--show` window) and exit. The window stays off screen. Text passed to "insert into reply" is written to `out.png.inserted.txt`. |
| `--outlook` | Work on the email selected in the running Outlook instead of the sample thread. |
| `--dump file` | Read the email selected in the running Outlook, write the shape of its thread to the file, and exit. Only counts, sizes, dates and flags are written, never names or text. |
| `--header file` | Work out where the reading-pane Assist button would go in the running Outlook, write the result to the file, and exit. Nothing is added to Outlook. |

```powershell
$exe = 'src\ProjectPA.DevHost\bin\Release\net48\ProjectPA.DevHost.exe'
Start-Process $exe                                                       # interactive window
Start-Process $exe -Wait -ArgumentList '--model haiku --do draft --click Shorter --shot out.png'
Start-Process $exe -Wait -ArgumentList '--dump thread-shape.txt'         # checks the Outlook-reading code
```

It is a windowed program, so use `Start-Process -Wait` when a script needs to wait for it.

`--dump` and `--outlook` attach to the running Outlook from outside. That exercises `OutlookHost`, the class with all the Outlook object-model code, without installing a new build or restarting Outlook. Outlook allows this without a prompt as long as Windows reports an active, up-to-date antivirus.

## Checking the ribbon

A mistake in the ribbon XML makes Outlook drop the whole ribbon silently. The installed add-in can be asked for the XML it would give each window, which catches malformed XML and duplicate ids before Outlook is involved:

```powershell
$c = New-Object -ComObject ProjectPA.Connect
foreach ($id in 'Microsoft.Outlook.Explorer', 'Microsoft.Outlook.Mail.Read', 'Microsoft.Outlook.Mail.Compose') {
    $doc = [xml]$c.GetCustomUI($id)
    $ids = $doc.SelectNodes('//*[@id]') | ForEach-Object { $_.GetAttribute('id') }
    "$id : $($ids.Count) ids, duplicates: $(($ids | Group-Object | Where-Object Count -gt 1).Name)"
}
```

To see Outlook's own ribbon errors, turn on File, Options, Advanced, Developers, "Show add-in user interface errors".

## Runtime data

Everything the add-in writes at run time is under `%LOCALAPPDATA%\ProjectPA`:

| Path | Contents |
|---|---|
| `app\<timestamp>` | Installed builds. |
| `settings.json` | Model, effort and other options. |
| `logs\<date>.log` | One log file per day. |
| `sessions\<timestamp>` | One folder per email worked on: `system.md` (the system prompt), `thread.md` (the thread as sent to Claude), `attachments\`, and `session.id` (the Claude Code conversation, once there is one). Deleted after the number of days set in Settings. |
| `prompts\<name>.md` | Prompts the user changed in the Prompts window. A file here replaces the built-in prompt of the same name. |
| `style\<account>.md` | The writing-style description learned for an account, added to every draft request for it. |

Setting the environment variable `PROJECTPA_DATA` to a folder makes everything above live there instead. The unit tests use it so that they never touch real settings or prompts, and it is handy for trying something in the DevHost without disturbing the installed add-in:

```powershell
$env:PROJECTPA_DATA = "$env:TEMP\papii-scratch"
Start-Process src\ProjectPA.DevHost\bin\Release\net48\ProjectPA.DevHost.exe
```

## Troubleshooting

**The PApii tab does not appear.** Open File, Options, Add-ins. If PApii is listed under Inactive or Disabled, select "COM Add-ins" (or "Disabled Items") in the Manage box and re-enable it. Check `LoadBehavior` in the registry key above: 3 means load at startup; Outlook sets it to 2 after a load failure. Then read the newest file in `%LOCALAPPDATA%\ProjectPA\logs`; a successful load writes a `connected` line.

**"Claude Code was not found."** The add-in looks for `claude.exe` in this order: the path in `settings.json` (`ClaudePath`), the `PATH`, `%USERPROFILE%\.local\bin`, and then the newest copy bundled with the Claude Code extension for VS Code. Set `ClaudePath` if yours is elsewhere.

**Claude reports that you are not logged in.** Run `claude` once in a terminal and sign in. The add-in uses the same login. It removes `ANTHROPIC_API_KEY` from the environment of the process it starts, so that a key set on the machine does not divert usage from your subscription to API billing.

**Requests are slow.** Each successful request writes a line to the log with the total time, the time spent inside the API, and the wait for the first word. If the API time accounts for nearly all of it, the delay is on the service side (it tends to grow as the plan's usage limit is approached) and not in the add-in.

**A build fails with a file-in-use error.** The source folder is inside OneDrive, and sync can briefly lock build output. Build again.

## Git

On the development machine git is available only in WSL, in the `Ubuntu-20.04` distribution, and the remote uses SSH (`git@github.com:bfjia/ProjectPA.git`). From PowerShell:

```powershell
$w = '/mnt/d/OneDrive/ProjectMisc/ProjectPA'
wsl.exe -d Ubuntu-20.04 --cd $w -e git status
wsl.exe -d Ubuntu-20.04 --cd $w -e git add -A
wsl.exe -d Ubuntu-20.04 --cd $w -e git commit -m 'Message'
wsl.exe -d Ubuntu-20.04 --cd $w -e git push
```

The drive is mounted without Unix permission metadata, so the repository sets `core.fileMode false`.

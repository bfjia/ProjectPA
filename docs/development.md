# Development

This document covers building, testing, installing and debugging ProjectPA.

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
| `src/ProjectPA.Core` | Logic with no Outlook or UI dependency: the Claude runner, settings, logging, prompts. |
| `src/ProjectPA.UI` | The WPF assistant pane and the `Assistant` class that drives it. No Outlook dependency. |
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

```powershell
$exe = 'src\ProjectPA.DevHost\bin\Release\net48\ProjectPA.DevHost.exe'
& $exe                                            # interactive window
& $exe --ask "Draft a short thank-you note"       # send a prompt on start
& $exe --model haiku --ask "..." --shot out.png   # run, save a picture of the pane, exit
```

With `--shot` the window is kept off screen and the process exits when the reply is complete, which makes it usable from scripts. `--model` applies to that run only and is not saved.

## Runtime data

Everything the add-in writes at run time is under `%LOCALAPPDATA%\ProjectPA`:

| Path | Contents |
|---|---|
| `app\<timestamp>` | Installed builds. |
| `settings.json` | Model, effort and other options. |
| `logs\<date>.log` | One log file per day. |
| `sessions\<timestamp>` | One folder per assistant conversation: the system prompt and, from Phase 1, the email thread and its attachments. |
| `prompts\<name>.md` | Optional. A file here replaces the built-in prompt of the same name. |

## Troubleshooting

**The Assistant tab does not appear.** Open File, Options, Add-ins. If ProjectPA Assistant is listed under Inactive or Disabled, select "COM Add-ins" (or "Disabled Items") in the Manage box and re-enable it. Check `LoadBehavior` in the registry key above: 3 means load at startup; Outlook sets it to 2 after a load failure. Then read the newest file in `%LOCALAPPDATA%\ProjectPA\logs`; a successful load writes a `connected` line.

**"Claude Code was not found."** The add-in looks for `claude.exe` in this order: the path in `settings.json` (`ClaudePath`), the `PATH`, `%USERPROFILE%\.local\bin`, and then the newest copy bundled with the Claude Code extension for VS Code. Set `ClaudePath` if yours is elsewhere.

**Claude reports that you are not logged in.** Run `claude` once in a terminal and sign in. The add-in uses the same login. It removes `ANTHROPIC_API_KEY` from the environment of the process it starts, so that a key set on the machine does not divert usage from your subscription to API billing.

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

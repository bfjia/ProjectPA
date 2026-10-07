# ProjectPA

ProjectPA builds **PApii**, an add-in for classic Outlook on Windows that acts as a personal assistant for email. It drafts replies from the full email thread and its attachments, summarises threads, tells you what an email needs from you, proposes meeting times that are free in your calendar, and adds agreed meetings to it. Generation runs through the Claude Code CLI using your own Claude subscription.

Every action is triggered by you. PApii never sends an email, creates a calendar item, or changes anything in Outlook without an explicit click.

"ProjectPA" is the name of this repository and of the code. "PApii" is the name you see in Outlook: the ribbon tab, the pane and the add-in entry.

## Status

In development. This file and the documents in `docs/` are updated as each feature lands.

| Phase | Scope | State |
|---|---|---|
| 0 | Repository, solution skeleton, add-in registration, pane, Claude integration | Done |
| 1 | Reply drafting, Assist, Summarize, follow-up questions, ribbon options, settings, prompt editor | Done |
| 2 | Find meeting times, tentative holds, add to calendar with a time-zone choice | Done |
| 3 | Write from a brief, polish your own text, extract tasks, follow-up reminders, writing-style learning | Done |
| 4 | Saved sessions: listed on the ribbon, removed automatically after a week, or all at once | Done |

## Future work

Not started. Listed here so they are not lost.

- **Send invitations from Add to Calendar.** A toggle on the event card that sends a meeting request to the other participants, off by default. Today PApii only adds the entry to your own calendar.
- **Filling in and signing attached forms.** A button that starts a guided conversation to fill a PDF or Word form from the thread and your own details, stamp a signature image where you approve it, and attach the result to a reply. This needs a document library and is the one feature expected to require installing something extra.
- **Custom ribbon icons.** The buttons use icons built into Office, picked for the closest meaning.
- **Start-up and response-time tuning.**

## Dependencies

### To run PApii

| Dependency | Version | Notes |
|---|---|---|
| Windows | 10 or 11, 64-bit | |
| Classic Outlook | Microsoft 365 or 2016 and later, **64-bit** | Developed and tested on Microsoft 365, version 16.0, 64-bit. The "new Outlook" cannot load this kind of add-in. The install script registers for 64-bit Outlook only. |
| .NET Framework | 4.8 or later | Part of Windows 10 (version 1903 onwards) and Windows 11. Nothing to install. |
| Claude Code | Tested with 2.1.292 | Signed in with a Claude subscription (Pro or Max). Either the standalone CLI or the Claude Code extension for VS Code, which includes the CLI. |
| Internet access | | Requests go to Anthropic through Claude Code. |

### To build it

| Dependency | Version | Notes |
|---|---|---|
| .NET SDK | 9.0 (developed with 9.0.201) | Provides `dotnet build` and `dotnet test`. |
| .NET Framework 4.8 reference assemblies | | Present when Visual Studio or the .NET Framework 4.8 Developer Pack is installed. Otherwise the .NET SDK fetches them from NuGet on the first build. |
| Office interop assemblies | `Microsoft.Office.Interop.Outlook` 15.0, `office` 15.0, `Extensibility` 7.0 | Installed with Office into `C:\Windows\assembly`. Referenced by path in `src/ProjectPA.AddIn/ProjectPA.AddIn.csproj`; the types are embedded, so no interop DLL is shipped. |
| Git | any | Only to clone the repository. |

NuGet packages, restored automatically from nuget.org on the first build:

| Package | Version | Used by |
|---|---|---|
| `Newtonsoft.Json` | 13.0.3 | The add-in (the only library shipped with it) |
| `xunit` | 2.9.2 | Tests |
| `xunit.runner.visualstudio` | 2.8.2 | Tests |
| `Microsoft.NET.Test.Sdk` | 17.11.1 | Tests |

**Not needed:** Visual Studio, the Office/VSTO developer tools, administrator rights, Node.js, Python.

## Install

Run the commands in Windows PowerShell from the repository folder.

**1. Install and sign in to Claude Code**, if you have not already. Follow the [Claude Code quickstart](https://code.claude.com/docs/en/quickstart), or install the [Claude Code extension for VS Code](https://code.claude.com/docs/en/vs-code). Start it once and sign in with your Claude account. To check:

```powershell
claude auth status     # should show "loggedIn": true and your subscription type
```

If `claude` is not on your `PATH` because you only have the VS Code extension, that is fine: PApii finds the copy inside the extension by itself.

**2. Check the build tools.**

```powershell
dotnet --list-sdks                                                        # expect a 9.0 line
Test-Path C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Outlook   # expect True
```

**3. Get the code.**

```powershell
git clone https://github.com/bfjia/ProjectPA.git
cd ProjectPA
```

**4. Build and run the tests.**

```powershell
scripts\build.ps1
```

**5. Install into Outlook.**

```powershell
scripts\install.ps1
```

The script first checks what it depends on: 64-bit PowerShell, 64-bit classic Outlook, .NET Framework 4.8, a finished build, and Claude Code installed and signed in. If anything is missing it lists each problem with how to fix it, changes nothing, and stops. `scripts\install.ps1 -SkipChecks` installs regardless.

This needs no administrator rights. It copies the build to `%LOCALAPPDATA%\ProjectPA\app\<timestamp>` and registers the add-in for your user under `HKEY_CURRENT_USER`. It does not close or restart Outlook.

**6. Restart Outlook.** A **PApii** tab appears in the ribbon, and a PApii group at the end of the Home tab. Select an email and click **Draft Reply**.

If PowerShell refuses to run the scripts because of its execution policy, run them as `powershell -ExecutionPolicy Bypass -File scripts\build.ps1` (and the same for `install.ps1`).

### Updating

Pull the new code, then repeat steps 4 to 6. Installing works while Outlook is open; the new build is loaded at the next Outlook start.

### Uninstalling

```powershell
scripts\uninstall.ps1              # removes the registration and the installed files
scripts\uninstall.ps1 -PurgeData   # also deletes settings, your edited prompts, logs and saved sessions
```

Close Outlook first so the files can be deleted.

### If the PApii tab does not appear

Open File, Options, Add-ins in Outlook and look for PApii under Inactive or Disabled. The troubleshooting section of [docs/development.md](docs/development.md) covers this and other problems.

## Documentation

- [docs/user-guide.md](docs/user-guide.md): every button, flow and setting, and what data is sent where.
- [docs/architecture.md](docs/architecture.md): how the add-in works and why it is built this way.
- [docs/development.md](docs/development.md): building, testing, installing and troubleshooting.

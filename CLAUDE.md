# ProjectPA

A COM add-in for classic Outlook that acts as a personal assistant for email. It appears in Outlook as **PApii**; "ProjectPA" is the name of the repo, the namespaces, the data folder and the registry ids. Generation runs through the Claude Code CLI on the owner's Claude subscription. Every action is triggered by a click; the add-in never sends mail or changes Outlook on its own.

Read `docs/architecture.md` before changing how anything works. `docs/user-guide.md` describes every control, and `docs/development.md` covers build, install and troubleshooting.

## Working rules (set by the owner)

- **Ask before installing any tooling.** NuGet packages already in use are approved: Newtonsoft.Json, xunit, xunit.runner.visualstudio, Microsoft.NET.Test.Sdk.
- **Never close, restart or kill Outlook.** After installing a build, say a restart is needed and wait.
- **Work in this folder only.** Agreed exceptions: `%LOCALAPPDATA%\ProjectPA` (installed builds and runtime data) and HKCU registry keys for add-in registration.
- **Commit and push as you go** when a change is significant.
- **Code is terse ("caveman" style); documentation is full prose.** Fewest files, types and lines that do the job; no speculative interfaces; comments rare and short, only for non-obvious reasons. Docs are complete sentences and are updated in the same commit as the feature.
- **User-facing name is PApii**, never "Assistant". The phrase "personal assistant" is fine as a description. The "Assist" button keeps its name.

## Environment

- Windows 11, classic Outlook 64-bit (Microsoft 365). Accounts: two Exchange, one Gmail over IMAP.
- .NET SDK 9, targeting .NET Framework 4.8. Office interop assemblies come from the GAC (`C:\Windows\assembly`) and are embedded.
- Visual Studio has no Office/VSTO workload; do not rely on VSTO.
- Claude Code on Windows exists only inside the VS Code extension (`%USERPROFILE%\.vscode\extensions\anthropic.claude-code-*\resources\native-binary\claude.exe`). `Claude.Find` discovers it.
- The owner's plan is Claude Pro. Usage limits are real: see "Testing" below.

### Git

Git is not on Windows. Use WSL distro `Ubuntu-20.04` over SSH:

```powershell
$w = '/mnt/d/OneDrive/ProjectMisc/ProjectPA'
wsl.exe -d Ubuntu-20.04 --cd $w -e git add -A
wsl.exe -d Ubuntu-20.04 --cd $w -e git commit -q -m 'Message'
wsl.exe -d Ubuntu-20.04 --cd $w -e git push -q
```

Remote is `git@github.com:bfjia/ProjectPA.git` (HTTPS has no credentials). Git there is 2.25: no `init -b`. `core.fileMode` is false. Avoid double quotes inside `wsl.exe -e bash -lc '...'` from PowerShell 5.1; use `git commit -F <file>` for multi-line messages.

## Layout

| Path | Contents |
|---|---|
| `src/ProjectPA.Core` | No Outlook, no UI. `Claude.cs` (CLI runner), `Context.cs` (thread rendering, quote stripping, attachment text, META parsing, sessions), `Scheduling.cs` (free windows, time zones, JSON schemas), `Email.cs` (models and `IHost`), `Util.cs` (Paths, Log, Settings, Prompts), `Prompts/*.md` |
| `src/ProjectPA.UI` | WPF. `PApii.cs` (pane state and every action), `PApiiPane.xaml`, `SettingsWindow`, `PromptsWindow` |
| `src/ProjectPA.AddIn` | `Connect.cs` (COM entry point, ribbon callbacks, pane hosting), `OutlookHost.cs` (all Outlook object-model code, implements `IHost`), `Ribbon.xml` |
| `src/ProjectPA.DevHost` | Runs the pane in a plain window on a sample thread; `SampleHost` fakes Outlook |
| `tests/ProjectPA.Core.Tests` | xunit tests for Core |
| `scripts` | `build.ps1`, `install.ps1`, `uninstall.ps1` |

## Build, test, install

```powershell
scripts\build.ps1      # dotnet build -c Release, then dotnet test
scripts\install.ps1    # copies to %LOCALAPPDATA%\ProjectPA\app\<timestamp>, registers under HKCU; never touches Outlook
```

The new build loads at the next Outlook start. Before telling the owner to restart, pre-flight from 64-bit PowerShell:

```powershell
$c = New-Object -ComObject ProjectPA.Connect
foreach ($id in 'Microsoft.Outlook.Explorer','Microsoft.Outlook.Mail.Read','Microsoft.Outlook.Mail.Compose') { [xml]$c.GetCustomUI($id) | Out-Null }
New-Object -ComObject ProjectPA.PaneHost | Out-Null
```

Check ribbon ids for duplicates too (`docs/development.md` has the snippet). A malformed ribbon makes Outlook drop it silently.

## How it works, in brief

- **Add-in type.** Plain COM add-in (`IDTExtensibility2`, `IRibbonExtensibility`, `ICustomTaskPaneConsumer`), not VSTO and not Office.js (web add-ins do not load on IMAP accounts). The task pane is a COM-visible WinForms control hosting WPF through `ElementHost`. WebView2 was ruled out for known focus bugs in Office panes.
- **Claude.** `Claude.Run` starts `claude -p` per request with `--restricted --tools Read --permission-mode dontAsk --strict-mcp-config --disable-slash-commands --system-prompt-file system.md`, streaming `stream-json`. The first turn uses `--session-id`, later turns `--resume`. Requests for data use `--output-format json --json-schema` and read `structured_output`. `ANTHROPIC_API_KEY` is removed from the child environment so the subscription is used. Do not use `--bare`: it ignores the subscription login.
- **Sessions.** One folder per email under `%LOCALAPPDATA%\ProjectPA\sessions` with `thread.md`, `system.md`, `attachments\`. The thread is sent inline in the first prompt. Sessions older than `KeepDays` (7) are removed at Outlook start; removal also runs `claude purge <folder> --yes`.
- **Reply format.** Prose first, then a line `---META---`, then one line of JSON (alternative intents, agreed meeting, suggested actions). `Context.Visible` hides it while streaming; `Context.SplitMeta` separates it.
- **Host boundary.** The pane only talks to `IHost`. Anything touching Outlook goes in `OutlookHost`; add the member to `IHost` and to `SampleHost`.
- **Threading.** Every action enters through `PApii.Go`, which runs it on the pane's dispatcher. Ribbon callbacks have no sync context.
- **Scheduling.** Calendar reading and free-window maths are local (`Scheduling.FreeWindows`, unit tested). Claude gets only free windows, never calendar entries, and every time it returns is re-checked against the calendar.
- **Settings.** `Settings.Current` re-reads `settings.json` when the file timestamp changes. Defaults: Opus, medium effort. Deserialization uses `ObjectCreationHandling.Replace` so list defaults are not duplicated.
- **Prompts.** Embedded files in `Core/Prompts`; a file of the same name in `%LOCALAPPDATA%\ProjectPA\prompts` overrides. `Prompts.All` lists what the Prompts window edits. The learned writing style is appended to the system prompt in code.

## Gotchas learned the hard way

- WPF projects on net48 do not get `System.IO` as an implicit using; add it. `System.Net.Http` is removed in `Directory.Build.targets`.
- Embedded interop returns `dynamic` for untyped members (`Selection[1]`, `CurrentItem`, `Context`, `ContentControl`). Cast to `(object)` first to keep binding static.
- The message context menu (`ContextMenuMailItem`) exists only in the Explorer ribbon; `GetCustomUI` strips it for the other two.
- Ribbon icons: use only `imageMso` names already in `Ribbon.xml`, which are confirmed to display. `SchedulingAssistant` is not a valid name. There is no schema file on this machine to validate against.
- Outlook `Restrict` date filters use the current culture's `g` format; recurring items need `Sort("[Start]")` and `IncludeRecurrences = true` before a bounded filter.
- IMAP mailboxes cannot hold dated flags; `FollowUp` creates a task for them.
- PowerShell: `$home` is reserved; a command combining `Remove-Item` with strings like `'/Type'` or `'~'` gets blocked by the tool's safety filter.
- Building in this OneDrive folder occasionally hits a file lock; build again.

## Testing

Live Claude calls spend the owner's own five-hour limit, shared with the coding session. Calls slow sharply as the limit is approached.

- Prefer unit tests and DevHost modes that make no Claude call: `--show settings|prompts`, `--dump`, `--do follow`.
- For live checks: `--model haiku`, one process at a time, one call per flow.
- DevHost keeps its data in `%TEMP%\ProjectPA-DevHost`. Never pass `--real-data` for a test: an earlier run saved its test model into the owner's real settings.
- `--dump file` attaches to the running Outlook and writes only counts, sizes and flags. Keep it that way; do not print names, subjects or message text from the owner's mailbox.
- DevHost is a windowed exe: run it with `Start-Process ... -PassThru` and `WaitForExit`. `--shot out.png` renders off screen for visual checks.
- Timings per request are in `%LOCALAPPDATA%\ProjectPA\logs` ("claude ok: ... ms in the API").

Nothing that writes to Outlook can be verified without the owner: say so plainly, and list what to try after a restart.

## Status

Done and confirmed working by the owner in Outlook: Draft Reply (with instructions, refinements, alternative intents), Assist, Summarize, follow-up questions, Find Times with tentative holds, Add to Calendar, Write from Brief, Polish, Extract Tasks, Follow Up, writing-style learning, prompt editor, settings, bold section labels.

Built in the latest round and not yet confirmed in Outlook: the Time zone menu on the event card (creating an event in another zone goes through `StartTimeZone` / `StartInStartTimeZone`), Saved Sessions menu with delete-all, clean-up at start, large Extract Tasks and Follow Up buttons, the Find Times icon.

Tried and removed (do not rebuild unasked): an Assist button injected into the reading pane header, and a button that opened a session in interactive Claude Code.

## Future work

- A toggle on the event card to send meeting invitations (off by default).
- Filling in and signing attached forms. Needs a document library; ask before installing (the earlier proposal was a private Python venv with PyMuPDF, python-docx, openpyxl).
- Custom ribbon icons.
- Start-up and response-time tuning.

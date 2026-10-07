# ProjectPA

A personal-assistant add-in for classic Outlook on Windows. It drafts replies from the full email thread, proposes meeting times from your calendar, adds confirmed meetings to the calendar, and handles other email admin. Generation runs through the Claude Code CLI using your own Claude subscription.

Every action is triggered by you. The add-in never sends an email, creates a calendar item, or changes anything in Outlook without an explicit click.

## Status

Early development. See the phases below; this file and the documents in `docs/` are updated as each feature lands.

| Phase | Scope | State |
|---|---|---|
| 0 | Repository, solution skeleton, add-in loads in Outlook, pane streams a reply from Claude | In progress |
| 1 | Reply drafting, Assist, Summarize, ribbon options, settings | Planned |
| 2 | Find meeting times, add to calendar | Planned |
| 3 | Compose and polish, tasks and follow-ups, writing-style learning | Planned |
| 4 | History, hand-off to Claude Code, polish | Planned |

Form filling and signing is planned for later and is not part of the current work.

## Requirements

- Windows with classic Outlook (Microsoft 365, 64-bit).
- .NET Framework 4.8 (included with Windows 10 and 11) and the .NET SDK to build.
- Claude Code installed and signed in with a Claude subscription. The add-in looks for `claude` on `PATH`, then `%USERPROFILE%\.local\bin\claude.exe`, then the copy bundled with the Claude Code extension for VS Code.

## Documentation

- [docs/architecture.md](docs/architecture.md): how the add-in works and why it is built this way.
- [docs/development.md](docs/development.md): building, testing, installing and troubleshooting.
- `docs/user-guide.md`: every button, flow and setting. Added with Phase 1.

## Quick start

```powershell
scripts\build.ps1      # build and run the tests
scripts\install.ps1    # install for the current user, then restart Outlook
```

An **Assistant** tab appears in the Outlook ribbon. To remove the add-in, run `scripts\uninstall.ps1`.

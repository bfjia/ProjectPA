# Architecture

ProjectPA is the name of this repository and of the code. The add-in it builds appears in Outlook as **PApii**.

PApii is a COM add-in for classic Outlook, written in C# on .NET Framework 4.8. It runs inside the Outlook process, reads mail and calendar data through the Outlook object model, and generates text by starting the Claude Code command-line tool as a child process.

## Why it is built this way

**A COM add-in rather than an Office web add-in.** Web add-ins (Office.js) only work on Exchange and Microsoft 365 mailboxes. One of the accounts this tool must serve is a Gmail account connected over IMAP, where web add-ins do not load. A COM add-in works on every account type and has full access to conversations, attachments and calendars.

**Plain COM rather than VSTO.** VSTO would need the Office development workload for Visual Studio, a multi-gigabyte install. Implementing the three COM interfaces directly needs no extra tooling and builds with `dotnet build`. The cost is that the add-in runs in Outlook's default application domain with no configuration file of its own, which is why dependencies are kept to a single library (Newtonsoft.Json).

**The Claude Code CLI rather than the API.** The requirement is to use a Claude subscription, and the CLI in headless mode (`claude -p`) is the supported way to do that from another program. No API key is involved.

**WPF for the pane rather than an embedded browser.** WebView2 has an open defect where it keeps keyboard focus inside Office task panes, and a reported failure in Outlook compose windows. WPF hosted through `ElementHost` avoids both.

## Projects

```
ProjectPA.Core     Claude runner, thread model and rendering, prompts, settings, logging   (no Outlook, no UI)
      ^
ProjectPA.UI       WPF pane, settings and prompts windows, and the PApii class behind the pane   (no Outlook)
      ^
ProjectPA.AddIn    COM add-in: ribbon, task pane hosting, all Outlook object-model code
      ^
ProjectPA.DevHost  the same pane in a plain window, on a sample thread or the running Outlook
```

Keeping Outlook out of the core and UI projects means almost everything can be exercised by unit tests and by the DevHost, without restarting Outlook. The pane's own logic is unit tested as well: `PApii.Runner` is the function a request to Claude goes through, `Claude.Run` by default, and the tests replace it with one that returns a prepared answer, next to a made-up `IHost`.

## How Outlook loads the add-in

Outlook reads `HKCU\Software\Microsoft\Office\Outlook\Addins\ProjectPA.Connect` at startup and creates the class registered under that name. `Connect` (in `src/ProjectPA.AddIn/Connect.cs`) implements three interfaces:

| Interface | What Outlook uses it for |
|---|---|
| `IDTExtensibility2` | Lifecycle. `OnConnection` hands over the Outlook `Application` object. It does almost nothing, because Outlook disables add-ins that start slowly. |
| `IRibbonExtensibility` | `GetCustomUI` returns the ribbon XML (`Ribbon.xml`) for the main window, the read window and the compose window. The same file serves all three: the name of each window's first built-in tab is substituted in, and the message context menu is removed for the two windows that do not have one, because a ribbon that names a control missing from its window fails to load at all. Ribbon callbacks named in the XML are public methods on `Connect`, found by name; each control's `tag` says which action it triggers. |
| `ICustomTaskPaneConsumer` | Outlook passes a factory for creating task panes. |

A task pane must be an ActiveX control, so `PaneHost` is a COM-visible Windows Forms control that contains an `ElementHost`, which in turn contains the WPF `PApiiPane`. Each Outlook window gets its own pane, created the first time a PApii button is clicked in that window. Panes are tracked by window handle, and entries for closed windows are dropped the next time a pane is requested.

`Connect` also installs an assembly-resolve handler. Outlook loads the add-in from its install folder, but WPF looks up assemblies by name, which does not search that folder; the handler fills the gap.

## Tried and removed

Two things were built and then taken out. They are recorded here so they are not rebuilt by accident.

**A button in the reading pane header.** Outlook gives COM add-ins no way to put a button next to Reply, Reply All and Forward. An Assist button was placed there anyway, as a child window of the header (an ordinary dialog, the same one Copilot's Summarize button lives in), kept in position by a timer. It worked but did not feel solid in use, and it depended on Outlook internals. The ribbon, the Home tab group and the right-click menu are the supported places and are the only ones now.

**Hand-off to interactive Claude Code.** A button opened a terminal running Claude Code in a session's folder, continuing the pane's conversation with full tools. It was removed as not wanted. The Saved Sessions menu that grew out of it remains, and now opens the session's folder instead.

## The host interface

The pane never talks to Outlook directly. `IHost` (in `src/ProjectPA.Core/Email.cs`) is the small set of things PApii needs from a mail client: which item is current, read its thread, insert text into a reply, and list the accounts. `OutlookHost` in the add-in implements it with the Outlook object model, one instance per window. `SampleHost` in the DevHost implements it with a built-in sample thread. This is what lets all of PApii run, with real Claude calls, outside Outlook.

`IHost.CurrentKey` identifies the email an action in that window is about. For a received message it is the message's entry id. While a reply is being typed inline in the main window, it is still the id of the message being answered, which stays selected; so starting a reply does not change the key, and the pane keeps its session. Only a reply in its own window, which has no selected message behind it, gets a key of its own.

**A draft goes only to the email it was written for.** The Reply, Reply All and Insert buttons on a draft compare `CurrentKey` with the key of the thread the pane loaded (`PApii.Reply`). If another email has been selected in the meantime, nothing is inserted and the pane says which email the draft belongs to. Without this check the buttons would open a reply to whatever happened to be selected, addressed to the wrong people.

## How a request is put together

When a PApii button is clicked, `PApii.Act` (in `src/ProjectPA.UI/PApii.cs`) asks the host for the key of the current item. If it differs from the thread already loaded in the pane, the thread is loaded:

1. **Collecting messages.** `OutlookHost.ReadThread` takes the selected message, or the reply being written, and asks Outlook for its conversation. A conversation spans folders within the mailbox, so the user's own replies in Sent Items are included. Messages are de-duplicated (a message filed in two folders appears twice) and sorted oldest first. Each is marked as the user's own if it sits in Sent Items or was sent from the account's address. When a reply being written has no stored conversation yet, the quoted text inside it is used as the history.
2. **Attachments.** Files are saved to the session's `attachments` folder. Pictures embedded in the message body are skipped when smaller than 100 KB, because those are logos and signature images, and kept when larger, because those are usually pasted screenshots. Files over 15 MB and non-file attachments are skipped. Every skipped item is recorded with the reason.
3. **Digesting.** `Context.Digest` extracts text from Word, Excel and PowerPoint files (they are zip archives of XML, read with the framework's own zip and XML classes) and from plain-text formats. That text goes inline. PDFs and images stay as files, since Claude reads those itself. Anything else is dropped and noted.
4. **Rendering.** `Context.Render` writes the thread as text: a header naming the mailbox and the user, then one section per message with sender, recipients, date, attachment list and body. Quoted history is cut from each message by `Context.StripQuotes`, because the earlier messages are already present in full; without this a long thread would repeat itself many times over. The first message and forwards are left whole, since their quoted content may exist nowhere else. Bodies are capped at 20,000 characters and the thread at 150,000, dropping the oldest messages first.

   A quote is only cut where the thread really has its content. As the messages are rendered, the text kept so far is collected in a reduced form (`Context.Squeeze`: letters and digits only, with image placeholders and bracketed links removed, so that re-wrapping, `>` marks and rewritten links make no difference). The quoted part of a reply is then read from the bottom up, and it is kept as far as its last line that the collected text does not contain. Two cases depend on this. Answers written between the quoted lines stay, together with the lines they answer. A quoted message that never reached this mailbox, such as a reply sent to someone else only, stays with its From and Sent lines, while the older history beneath it, which the thread has, is still cut. Header lines, "On ... wrote:" lines and lines with fewer than 15 letters and digits never count as new, so a very short answer on its own below the last longer one can still be lost. When the content of a quote differs from the original for another reason, for example a disclaimer added by the other side's mail system, more of the quote is kept than needed; the method errs on the side of repeating text and not of losing it.

The rendered thread is saved as `thread.md` in the session folder and sent as the first part of the first prompt, followed by the request (for example the contents of `Prompts/draft-reply.md`). Sending the thread inline, instead of having Claude open a file, saves a tool round trip and several seconds. Later requests in the same pane send only the request and resume the session.

Prompts are embedded text files in `src/ProjectPA.Core/Prompts`. A file with the same name in `%LOCALAPPDATA%\ProjectPA\prompts` overrides the built-in one, and the Prompts window (`PromptsWindow` in the UI project) is the editor for those files. `Prompts.All` lists the prompts the editor offers. `Prompts.Save` writes the user's version, except when the text is empty or identical to the built-in text; then it removes the override, so that a later improvement to the built-in prompt still reaches the user. Placeholders such as `{{tone}}` are filled in by `Prompts.Get` each time a prompt is used. The system prompt is the exception to "each time": Claude Code records it at the start of a session, so an edit to it takes effect with the next email.

## Reading the reply

A draft needs to stream to the screen as prose, but the pane also wants structured extras: alternative replies to offer, and whether a meeting was agreed. The prompts therefore ask for the email body first, then a line containing `---META---`, then one line of JSON. `Context.Visible` hides the marker and everything after it while text streams in, including a marker that has only half arrived. `Context.SplitMeta` separates the two parts at the end. If the JSON is missing or malformed the body is still used; the extras are simply absent.

`PApii.Decorate` turns the result into a card: insert and copy buttons, the fixed refinement pills, and pills built from the JSON. Drafts are shown in a plain editable text box. Everything else is shown read-only with bold runs chosen by `Context.Runs`: a short "Label:" at the start of a line, and any text the model wrapped in `**`. This is done on the display side, so it works whatever wording a prompt uses. A follow-up typed by the user goes through `Prompts/followup.md`, which tells Claude to return a draft with the marker when a draft was asked for and a plain answer otherwise; the presence of the marker decides which kind of card is shown.

## Scheduling

Find Times and Add to Calendar split the work between Claude and local code along one line: Claude reads language, the add-in owns the calendar.

**Find Times** (`PApii.FindTimes`):

1. `OutlookHost.BusyBlocks` reads the calendars ticked in Settings from now to the end of the look-ahead period. Recurring meetings are expanded. Entries marked Free or Working Elsewhere are ignored. Only start and end times are kept.
2. `Scheduling.FreeWindows` (pure code, unit tested) turns those into free windows per day: inside working hours, on working days, with the buffer kept around each meeting, and nothing sooner than two hours from now.
3. One request goes to Claude with the thread and the list of free windows, and a JSON schema for the answer: title, length, the times the other side proposed, three to five times to offer, and a note. Claude is the right tool for this part because the constraints are in prose ("afternoons would be better", "the week after next", a time given in another zone).
4. The add-in does not trust the answer blindly. A time Claude suggests is kept only if the whole meeting fits inside one of the free windows from step 2 (`Scheduling.Fits`), which holds it to the working hours, the buffer and the look-ahead period as well as to the calendar. The other side's proposals are shown either way, marked free or conflicting by `Scheduling.IsFree`. If one of them lies beyond the period that was read, the calendar is read again as far as that day first; otherwise a date three weeks out would be reported as free only because nothing had been read for it.
5. The ticked times become instructions for an ordinary Draft Reply request.

What reaches Anthropic for this feature is the thread and the free windows. Calendar subjects, attendees and bodies never leave the machine.

**Add to Calendar** (`PApii.AddToCalendar` and `PApii.ShowEvent`): one request with a schema returns whether a time was found, the title, start, end, location, notes, whether both sides confirmed, and the evidence: the sentence in the thread the date and time were taken from, quoted word for word with its author and the date of that message. The result is shown as editable fields. Nothing is written until the user clicks the button, at which point `OutlookHost.CreateEvent` adds an appointment to the chosen calendar. When a draft or briefing already reported an agreed meeting in its `---META---` line, the same card is opened from that data with no further request.

The evidence is shown on the card as a "From the thread:" line. Claude does the arithmetic behind a date ("next Thursday", a time given in another zone), and that is where it has gone wrong before; the quoted sentence lets the user check the result against the source at a glance.

The card says whether the calendar already has something at that time. That check belongs to the time it was made for, so when the date, the times or the zone have been edited, clicking the button first checks the calendar for the new time. With a conflict, the card shows the warning and nothing is created; a second click with the same times adds the entry anyway.

**Holds.** "Hold on calendar" creates tentative appointments in the category "PApii hold", with a line in the body naming the thread's subject. `RemoveHolds` finds them again by that category and line when the real meeting is added, and deletes them (to Deleted Items). On the event card, PApii's own holds are not counted as conflicts. In Find Times they count as busy like any other entry.

**Time zones.** Each prompt is told the user's time zone and the current time, to convert times stated in other zones to local time, and to resolve relative days ("Thursday") from the date of the message that mentions them. The event card adds a Time zone menu, defaulting to the computer's zone, that says which zone the entered date and times are in. For another zone, `EventDraft.TimeZoneId` is set and `OutlookHost.CreateEvent` assigns that zone to the appointment and writes the times as wall-clock times in it, so Outlook stores the zone with the entry. If Outlook does not recognise the zone name, the times are converted to local time with `Scheduling.ToLocal` instead.

Requests that return data use the same model and effort as everything else. An earlier version forced low effort for speed; it was reverted after it produced a wrong date, since a wrong calendar entry costs more than a few seconds.

## Writing, tasks and style

**Sessions about an email being written.** Write from Brief and Polish are not about the selected thread, so `PApii.Writing` starts a session of a different kind: the "thread" given to Claude is a short description of the email being written (recipient, subject, and what has been typed so far, with quoted history removed). The `standalone` flag marks such a session. While it is set, follow-ups stay with that session whatever is selected in the message list, and the draft cards offer Insert (through `IHost.Compose`, which fills the open email or opens a new one) and, after Polish on highlighted text, Replace selection. Any thread action from the ribbon ends it by loading the selected email again.

**Tasks.** Extract Tasks is a request with a schema returning titles, due dates and notes. Ticked items go to `IHost.CreateTask`. Follow Up involves no request at all: `OutlookHost.FollowUp` flags the message with a reminder, or creates a task when the mailbox is IMAP, which cannot store a dated flag.

**Writing style.** Learn in Settings calls `IHost.SentSamples` for about 30 recent sent emails (the user's own text only), sends them in a single one-off request with the `style` prompt, and saves the reply to `%LOCALAPPDATA%\ProjectPA\style\<account>.md`. `Prompts.StyleNote` appends that file to the system prompt of every session for the account. It is appended in code and not through a placeholder, so it keeps working when the user has replaced the Ground rules prompt with their own.

## Saved sessions

Every email worked on gets a session folder under `%LOCALAPPDATA%\ProjectPA\sessions`, named by timestamp, holding `thread.md` (the thread exactly as sent), `system.md` and the saved attachments. Claude Code separately keeps its own transcript of each conversation under `~/.claude/projects`, keyed by that folder's path.

- **Automatic clean-up.** `Connect.OnConnection` starts a background task that calls `Context.PurgeSessions` with the retention period from Settings (seven days by default). It runs off the main thread, so Outlook's start-up is not delayed.
- **Removal is complete.** `Context.RemoveSessions` first runs `claude purge <folder> --yes`, which deletes Claude Code's transcript for that folder, and then deletes the folder.
- **Folders deleted by hand are caught up with.** Deleting a session folder in File Explorer leaves Claude Code's transcript behind. The clean-up at start therefore also looks for such orphans (`Context.Orphans`): Claude Code names each transcript folder after the path of the session folder, with every character that is not a letter or digit replaced by a dash, so the transcripts of sessions whose folder no longer exists can be listed and purged. This relies on how Claude Code names its folders; if that changes, the list is simply empty.
- **The Saved Sessions menu** is a `dynamicMenu` whose content `Connect.GetHistory` rebuilds each time it drops down from `Context.RecentSessions`, which reads the subject from the first line of each folder's `thread.md`. Picking an entry opens the folder in File Explorer. The last entry deletes all sessions after a confirmation.
- A pane whose session folder has been deleted reloads the thread on its next action.

## Settings

`Settings.Current` re-reads `settings.json` whenever the file's timestamp changes. Without that, a process holding an older copy in memory would silently write it back over a change made elsewhere. The defaults are Opus at medium effort.

**The account switch must not fail open.** Settings can switch PApii off for an account, and three things keep that promise (`PApii.Allow`):

- A `settings.json` that exists but cannot be read does not quietly become the defaults, which would switch every account back on. `Settings.Broken` then holds the reason, no mail is read, and `Save` does nothing, so the unreadable file is not replaced by defaults either. Saving from the Settings window, where the user sees and sets the accounts, clears the state.
- The account of an email comes from the store it sits in. A shared mailbox, an online archive or a data file is no account's own store, so `OutlookHost.ReadThread` falls back to the account a reply would be sent from.
- If the account still cannot be told while some account is switched off, the email is not read.

A refused email leaves nothing behind: the session folder, with any attachments already saved to it, is deleted.

The DevHost uses a scratch data folder in `%TEMP%` unless started with `--real-data`. This is deliberate: a DevHost test run once saved its temporary model choice into the real settings file.

## How generation works

`Claude.Run` in `src/ProjectPA.Core/Claude.cs` starts one `claude` process per request:

```
claude -p --model <model> --effort <effort> --tools Read
       --permission-mode dontAsk --strict-mcp-config --restricted --disable-slash-commands
       --output-format stream-json --verbose --include-partial-messages
       --system-prompt-file system.md
       --session-id <guid>            (first turn)   or   --resume <guid>   (follow-up turns)
```

The working directory is a session folder under `%LOCALAPPDATA%\ProjectPA\sessions`. The prompt is written to standard input as UTF-8. Output arrives as one JSON object per line, and `Claude.Feed` handles the three kinds that matter:

| Line type | Use |
|---|---|
| `stream_event` with a `text_delta` | Appended to the card on screen as it arrives. A `message_start` clears the card, so that text written before a tool call is replaced by the final answer. |
| `rate_limit_event` | Shown in the pane footer as the share of the five-hour limit used. |
| `result` | The final text, the session identifier, and the error message if the run failed. |

For requests that need data rather than prose (for example extracting a meeting time), the runner passes `--output-format json --json-schema <schema>` instead, and reads the validated object from `structured_output`.

Follow-up requests in the same pane reuse the session with `--resume`, so Claude keeps the context of the thread and earlier drafts.

### Why these flags

Email is untrusted input: a message can contain text written to manipulate the model that reads it. The flags limit what such text could achieve.

- `--tools Read` together with `--restricted` leaves Claude with one tool, reading files, confined to the session folder. There is no shell, no web access and no file writing. The confinement was checked on Claude Code 2.1.292 with these exact flags: a Read of a file outside the working directory is refused with "--restricted confines the file tools to the working directory".
- `--strict-mcp-config` and `--disable-slash-commands` keep MCP servers and skills configured for normal Claude Code use out of these runs. `--restricted` also ignores user and project settings files.
- `--permission-mode dontAsk` denies anything that would otherwise prompt, since nobody is there to answer.
- `--system-prompt-file` replaces Claude Code's default coding-oriented system prompt with a short one for email work, which also tells the model to treat email content as material and not as instructions.

The last line of defence is the design itself: output only goes into the pane, and reaches Outlook only when the user clicks a button.

### Finding the CLI and choosing the login

`Claude.Find` looks for `claude.exe` in the path from settings, then the `PATH`, then `%USERPROFILE%\.local\bin`, then the newest copy inside the Claude Code extension for VS Code (whose folder name changes with every extension update). The child process inherits the user's environment minus `ANTHROPIC_API_KEY` and `ANTHROPIC_AUTH_TOKEN`, so the CLI uses the subscription login.

One thing to watch: the Claude Code documentation says `--bare`, a mode that ignores subscription logins, will become the default for `-p` in a future release. All flags are built in `Claude.Args`, so adapting to that is a one-place change.

More generally, Claude Code updates itself (the copy inside the VS Code extension changes with every extension update), and the add-in depends on its flags and on the shape of its output. So that a failure after an update is easy to recognise, `Connect.OnConnection` writes the version and path of the CLI in use to the log each time Outlook starts (`Claude.Version`).

## Threading

Outlook's object model and WPF both belong to Outlook's main thread. The Claude process is read on background threads. `PApii.Go` is the single entry point for every action: it runs the action through the pane's dispatcher, so that code after each `await` resumes on the main thread even when the call started in a ribbon callback, which has no synchronisation context. Text deltas from Claude are posted to the same dispatcher.

## Runtime data

All run-time files live under `%LOCALAPPDATA%\ProjectPA`, outside the source folder and outside OneDrive. The layout is listed in `docs/development.md`.

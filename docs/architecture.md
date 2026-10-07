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

Keeping Outlook out of the core and UI projects means almost everything can be exercised by unit tests and by the DevHost, without restarting Outlook.

## How Outlook loads the add-in

Outlook reads `HKCU\Software\Microsoft\Office\Outlook\Addins\ProjectPA.Connect` at startup and creates the class registered under that name. `Connect` (in `src/ProjectPA.AddIn/Connect.cs`) implements three interfaces:

| Interface | What Outlook uses it for |
|---|---|
| `IDTExtensibility2` | Lifecycle. `OnConnection` hands over the Outlook `Application` object. It does almost nothing, because Outlook disables add-ins that start slowly. |
| `IRibbonExtensibility` | `GetCustomUI` returns the ribbon XML (`Ribbon.xml`) for the main window, the read window and the compose window. The same file serves all three: the name of each window's first built-in tab is substituted in, and the message context menu is removed for the two windows that do not have one, because a ribbon that names a control missing from its window fails to load at all. Ribbon callbacks named in the XML are public methods on `Connect`, found by name; each control's `tag` says which action it triggers. |
| `ICustomTaskPaneConsumer` | Outlook passes a factory for creating task panes. |

A task pane must be an ActiveX control, so `PaneHost` is a COM-visible Windows Forms control that contains an `ElementHost`, which in turn contains the WPF `PApiiPane`. Each Outlook window gets its own pane, created the first time a PApii button is clicked in that window. Panes are tracked by window handle, and entries for closed windows are dropped the next time a pane is requested.

`Connect` also installs an assembly-resolve handler. Outlook loads the add-in from its install folder, but WPF looks up assemblies by name, which does not search that folder; the handler fills the gap.

## The reading pane button

The Assist button in the message header is the one part of PApii that does not go through an Outlook extension point, because there is none for that area. It relies on an observation: the reading pane header is an ordinary Windows dialog (window class `#32770`) whose children include a toolbar (`ToolbarWindow32`) holding Reply, Reply All and Forward. Copilot's own Summarize button is a child window of the same dialog.

`HeaderButton` (in `src/ProjectPA.AddIn/HeaderButton.cs`) is a small custom-drawn control that PApii creates as another child of that dialog. A timer in `Connect` runs three times a second and calls `HeaderButton.Place` for each main window, which:

1. finds the visible header dialog under the main window and its visible toolbar;
2. computes a rectangle the height of the toolbar, immediately to its left;
3. hides the button if that rectangle would overlap any other visible control in the header (which happens when the pane is narrow and the sender line reaches that far), and otherwise moves the button there.

Being a child window, the button moves, clips and hides with the header for free; the timer only has to notice layout changes. When Outlook destroys the header and builds a new one, the button's window is destroyed with it and is recreated in the new header on the next tick.

Because this depends on Outlook internals, it is built to fail quietly:

- The button is never a standard `Button` control. A standard button reports clicks to its parent window, which here is Outlook's dialog and could mistake the message for one of its own commands. `HeaderButton` handles its own mouse input, and sets `WS_EX_NOPARENTNOTIFY` so the dialog is not told about it at all.
- It cannot take keyboard focus, so clicking it does not pull focus out of the message list.
- No Outlook window is subclassed, resized or moved. PApii only reads positions and places its own window.
- If the header or toolbar is not found, the button is simply hidden. If anything throws, the timer stops for the rest of the Outlook session and the error is logged once.
- It can be switched off in Settings (`HeaderButton` in `settings.json`).

The DevHost option `--header` runs the locating step against the running Outlook from outside and reports where the button would go, without creating it.

## The host interface

The pane never talks to Outlook directly. `IHost` (in `src/ProjectPA.Core/Email.cs`) is the small set of things PApii needs from a mail client: which item is current, read its thread, insert text into a reply, and list the accounts. `OutlookHost` in the add-in implements it with the Outlook object model, one instance per window. `SampleHost` in the DevHost implements it with a built-in sample thread. This is what lets all of PApii run, with real Claude calls, outside Outlook.

## How a request is put together

When a PApii button is clicked, `PApii.Run` (in `src/ProjectPA.UI/PApii.cs`) asks the host for the key of the current item. If it differs from the thread already loaded in the pane, the thread is loaded:

1. **Collecting messages.** `OutlookHost.ReadThread` takes the selected message, or the reply being written, and asks Outlook for its conversation. A conversation spans folders within the mailbox, so the user's own replies in Sent Items are included. Messages are de-duplicated (a message filed in two folders appears twice) and sorted oldest first. Each is marked as the user's own if it sits in Sent Items or was sent from the account's address. When a reply being written has no stored conversation yet, the quoted text inside it is used as the history.
2. **Attachments.** Files are saved to the session's `attachments` folder. Pictures embedded in the message body are skipped when smaller than 100 KB, because those are logos and signature images, and kept when larger, because those are usually pasted screenshots. Files over 15 MB and non-file attachments are skipped. Every skipped item is recorded with the reason.
3. **Digesting.** `Context.Digest` extracts text from Word, Excel and PowerPoint files (they are zip archives of XML, read with the framework's own zip and XML classes) and from plain-text formats. That text goes inline. PDFs and images stay as files, since Claude reads those itself. Anything else is dropped and noted.
4. **Rendering.** `Context.Render` writes the thread as text: a header naming the mailbox and the user, then one section per message with sender, recipients, date, attachment list and body. Quoted history is cut from each message by `Context.StripQuotes`, because the earlier messages are already present in full; without this a long thread would repeat itself many times over. The first message and forwards are left whole, since their quoted content may exist nowhere else. Bodies are capped at 20,000 characters and the thread at 150,000, dropping the oldest messages first.

The rendered thread is saved as `thread.md` in the session folder and sent as the first part of the first prompt, followed by the request (for example the contents of `Prompts/draft-reply.md`). Sending the thread inline, instead of having Claude open a file, saves a tool round trip and several seconds. Later requests in the same pane send only the request and resume the session.

Prompts are embedded text files in `src/ProjectPA.Core/Prompts`. A file with the same name in `%LOCALAPPDATA%\ProjectPA\prompts` overrides the built-in one, and the Prompts window (`PromptsWindow` in the UI project) is the editor for those files. `Prompts.All` lists the prompts the editor offers. `Prompts.Save` writes the user's version, except when the text is empty or identical to the built-in text; then it removes the override, so that a later improvement to the built-in prompt still reaches the user. Placeholders such as `{{tone}}` are filled in by `Prompts.Get` each time a prompt is used. The system prompt is the exception to "each time": Claude Code records it at the start of a session, so an edit to it takes effect with the next email.

## Reading the reply

A draft needs to stream to the screen as prose, but the pane also wants structured extras: alternative replies to offer, and whether a meeting was agreed. The prompts therefore ask for the email body first, then a line containing `---META---`, then one line of JSON. `Context.Visible` hides the marker and everything after it while text streams in, including a marker that has only half arrived. `Context.SplitMeta` separates the two parts at the end. If the JSON is missing or malformed the body is still used; the extras are simply absent.

`PApii.Decorate` turns the result into a card: insert and copy buttons, the fixed refinement pills, and pills built from the JSON. Drafts are shown in a plain editable text box. Everything else is shown read-only with bold runs chosen by `Context.Runs`: a short "Label:" at the start of a line, and any text the model wrapped in `**`. This is done on the display side, so it works whatever wording a prompt uses. A follow-up typed by the user goes through `Prompts/followup.md`, which tells Claude to return a draft with the marker when a draft was asked for and a plain answer otherwise; the presence of the marker decides which kind of card is shown.

## Scheduling

Find Times and Add to Calendar split the work between Claude and local code along one line: Claude reads language, the add-in owns the calendar.

**Find Times** (`PApii.FindTimes`):

1. `OutlookHost.BusyBlocks` reads the calendars ticked in Settings for the look-ahead period. Recurring meetings are expanded. Entries marked Free or Working Elsewhere are ignored. Only start and end times are kept.
2. `Scheduling.FreeWindows` (pure code, unit tested) turns those into free windows per day: inside working hours, on working days, with the buffer kept around each meeting, and nothing sooner than two hours from now.
3. One request goes to Claude with the thread and the list of free windows, and a JSON schema for the answer: title, length, the times the other side proposed, three to five times to offer, and a note. Claude is the right tool for this part because the constraints are in prose ("afternoons would be better", "the week after next", a time given in another zone).
4. The add-in does not trust the answer blindly: every time Claude returns is checked against the busy blocks again with `Scheduling.IsFree`, and suggestions that are not free are dropped. The other side's proposals are shown either way, marked free or conflicting.
5. The ticked times become instructions for an ordinary Draft Reply request.

What reaches Anthropic for this feature is the thread and the free windows. Calendar subjects, attendees and bodies never leave the machine.

**Add to Calendar** (`PApii.AddToCalendar` and `PApii.ShowEvent`): one request with a schema returns whether a time was found, the title, start, end, location, notes and whether both sides confirmed. The result is shown as editable fields. Nothing is written until the user clicks the button, at which point `OutlookHost.CreateEvent` adds an appointment to the chosen calendar. When a draft or briefing already reported an agreed meeting in its `---META---` line, the same card is opened from that data with no further request.

**Holds.** "Hold on calendar" creates tentative appointments in the category "PApii hold", with a line in the body naming the thread's subject. `RemoveHolds` finds them again by that category and line when the real meeting is added, and deletes them (to Deleted Items). Holds made by PApii are not counted as conflicts for the meeting they were made for.

Requests that return data run at low effort regardless of the ribbon setting, because extraction needs little deliberation and the user is waiting.

Each prompt is told the user's time zone and the current time, and to resolve relative days ("Thursday") from the date of the message that mentions them.

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

- `--tools Read` together with `--restricted` leaves Claude with one tool, reading files, confined to the session folder. There is no shell, no web access and no file writing.
- `--strict-mcp-config` and `--disable-slash-commands` keep MCP servers and skills configured for normal Claude Code use out of these runs. `--restricted` also ignores user and project settings files.
- `--permission-mode dontAsk` denies anything that would otherwise prompt, since nobody is there to answer.
- `--system-prompt-file` replaces Claude Code's default coding-oriented system prompt with a short one for email work, which also tells the model to treat email content as material and not as instructions.

The last line of defence is the design itself: output only goes into the pane, and reaches Outlook only when the user clicks a button.

### Finding the CLI and choosing the login

`Claude.Find` looks for `claude.exe` in the path from settings, then the `PATH`, then `%USERPROFILE%\.local\bin`, then the newest copy inside the Claude Code extension for VS Code (whose folder name changes with every extension update). The child process inherits the user's environment minus `ANTHROPIC_API_KEY` and `ANTHROPIC_AUTH_TOKEN`, so the CLI uses the subscription login.

One thing to watch: the Claude Code documentation says `--bare`, a mode that ignores subscription logins, will become the default for `-p` in a future release. All flags are built in `Claude.Args`, so adapting to that is a one-place change.

## Threading

Outlook's object model and WPF both belong to Outlook's main thread. The Claude process is read on background threads. `PApii.Go` is the single entry point for every action: it runs the action through the pane's dispatcher, so that code after each `await` resumes on the main thread even when the call started in a ribbon callback, which has no synchronisation context. Text deltas from Claude are posted to the same dispatcher.

## Runtime data

All run-time files live under `%LOCALAPPDATA%\ProjectPA`, outside the source folder and outside OneDrive. The layout is listed in `docs/development.md`.

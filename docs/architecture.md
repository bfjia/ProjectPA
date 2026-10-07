# Architecture

ProjectPA is a COM add-in for classic Outlook, written in C# on .NET Framework 4.8. It runs inside the Outlook process, reads mail and calendar data through the Outlook object model, and generates text by starting the Claude Code command-line tool as a child process.

## Why it is built this way

**A COM add-in rather than an Office web add-in.** Web add-ins (Office.js) only work on Exchange and Microsoft 365 mailboxes. One of the accounts this tool must serve is a Gmail account connected over IMAP, where web add-ins do not load. A COM add-in works on every account type and has full access to conversations, attachments and calendars.

**Plain COM rather than VSTO.** VSTO would need the Office development workload for Visual Studio, a multi-gigabyte install. Implementing the three COM interfaces directly needs no extra tooling and builds with `dotnet build`. The cost is that the add-in runs in Outlook's default application domain with no configuration file of its own, which is why dependencies are kept to a single library (Newtonsoft.Json).

**The Claude Code CLI rather than the API.** The requirement is to use a Claude subscription, and the CLI in headless mode (`claude -p`) is the supported way to do that from another program. No API key is involved.

**WPF for the pane rather than an embedded browser.** WebView2 has an open defect where it keeps keyboard focus inside Office task panes, and a reported failure in Outlook compose windows. WPF hosted through `ElementHost` avoids both.

## Projects

```
ProjectPA.Core     Claude runner, thread model and rendering, prompts, settings, logging   (no Outlook, no UI)
      ^
ProjectPA.UI       WPF pane, settings window, and the Assistant class that drives them     (no Outlook)
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

A task pane must be an ActiveX control, so `PaneHost` is a COM-visible Windows Forms control that contains an `ElementHost`, which in turn contains the WPF `AssistantPane`. Each Outlook window gets its own pane, created the first time an assistant button is clicked in that window. Panes are tracked by window handle, and entries for closed windows are dropped the next time a pane is requested.

`Connect` also installs an assembly-resolve handler. Outlook loads the add-in from its install folder, but WPF looks up assemblies by name, which does not search that folder; the handler fills the gap.

## The host interface

The pane never talks to Outlook directly. `IHost` (in `src/ProjectPA.Core/Email.cs`) is the small set of things the assistant needs from a mail client: which item is current, read its thread, insert text into a reply, and list the accounts. `OutlookHost` in the add-in implements it with the Outlook object model, one instance per window. `SampleHost` in the DevHost implements it with a built-in sample thread. This is what lets the whole assistant run, with real Claude calls, outside Outlook.

## How a request is put together

When an assistant button is clicked, `Assistant.Run` (in `src/ProjectPA.UI/Assistant.cs`) asks the host for the key of the current item. If it differs from the thread already loaded in the pane, the thread is loaded:

1. **Collecting messages.** `OutlookHost.ReadThread` takes the selected message, or the reply being written, and asks Outlook for its conversation. A conversation spans folders within the mailbox, so the user's own replies in Sent Items are included. Messages are de-duplicated (a message filed in two folders appears twice) and sorted oldest first. Each is marked as the user's own if it sits in Sent Items or was sent from the account's address. When a reply being written has no stored conversation yet, the quoted text inside it is used as the history.
2. **Attachments.** Files are saved to the session's `attachments` folder. Pictures embedded in the message body are skipped when smaller than 100 KB, because those are logos and signature images, and kept when larger, because those are usually pasted screenshots. Files over 15 MB and non-file attachments are skipped. Every skipped item is recorded with the reason.
3. **Digesting.** `Context.Digest` extracts text from Word, Excel and PowerPoint files (they are zip archives of XML, read with the framework's own zip and XML classes) and from plain-text formats. That text goes inline. PDFs and images stay as files, since Claude reads those itself. Anything else is dropped and noted.
4. **Rendering.** `Context.Render` writes the thread as text: a header naming the mailbox and the user, then one section per message with sender, recipients, date, attachment list and body. Quoted history is cut from each message by `Context.StripQuotes`, because the earlier messages are already present in full; without this a long thread would repeat itself many times over. The first message and forwards are left whole, since their quoted content may exist nowhere else. Bodies are capped at 20,000 characters and the thread at 150,000, dropping the oldest messages first.

The rendered thread is saved as `thread.md` in the session folder and sent as the first part of the first prompt, followed by the request (for example the contents of `Prompts/draft-reply.md`). Sending the thread inline, instead of having Claude open a file, saves a tool round trip and several seconds. Later requests in the same pane send only the request and resume the session.

Prompts are embedded text files in `src/ProjectPA.Core/Prompts`. A file with the same name in `%LOCALAPPDATA%\ProjectPA\prompts` overrides the built-in one. Placeholders such as `{{tone}}` are filled in by `Prompts.Get`.

## Reading the reply

A draft needs to stream to the screen as prose, but the pane also wants structured extras: alternative replies to offer, and whether a meeting was agreed. The prompts therefore ask for the email body first, then a line containing `---META---`, then one line of JSON. `Context.Visible` hides the marker and everything after it while text streams in, including a marker that has only half arrived. `Context.SplitMeta` separates the two parts at the end. If the JSON is missing or malformed the body is still used; the extras are simply absent.

`Assistant.Decorate` turns the result into a card: insert and copy buttons, the fixed refinement pills, and pills built from the JSON. A follow-up typed by the user goes through `Prompts/followup.md`, which tells Claude to return a draft with the marker when a draft was asked for and a plain answer otherwise; the presence of the marker decides which kind of card is shown.

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

Email is untrusted input: a message can contain text written to manipulate an assistant. The flags limit what such text could achieve.

- `--tools Read` together with `--restricted` leaves Claude with one tool, reading files, confined to the session folder. There is no shell, no web access and no file writing.
- `--strict-mcp-config` and `--disable-slash-commands` keep MCP servers and skills configured for normal Claude Code use out of these runs. `--restricted` also ignores user and project settings files.
- `--permission-mode dontAsk` denies anything that would otherwise prompt, since nobody is there to answer.
- `--system-prompt-file` replaces Claude Code's default coding-oriented system prompt with a short one for email work, which also tells the model to treat email content as material and not as instructions.

The last line of defence is the design itself: output only goes into the pane, and reaches Outlook only when the user clicks a button.

### Finding the CLI and choosing the login

`Claude.Find` looks for `claude.exe` in the path from settings, then the `PATH`, then `%USERPROFILE%\.local\bin`, then the newest copy inside the Claude Code extension for VS Code (whose folder name changes with every extension update). The child process inherits the user's environment minus `ANTHROPIC_API_KEY` and `ANTHROPIC_AUTH_TOKEN`, so the CLI uses the subscription login.

One thing to watch: the Claude Code documentation says `--bare`, a mode that ignores subscription logins, will become the default for `-p` in a future release. All flags are built in `Claude.Args`, so adapting to that is a one-place change.

## Threading

Outlook's object model and WPF both belong to Outlook's main thread. The Claude process is read on background threads. `Assistant.Go` is the single entry point for every action: it runs the action through the pane's dispatcher, so that code after each `await` resumes on the main thread even when the call started in a ribbon callback, which has no synchronisation context. Text deltas from Claude are posted to the same dispatcher.

## Runtime data

All run-time files live under `%LOCALAPPDATA%\ProjectPA`, outside the source folder and outside OneDrive. The layout is listed in `docs/development.md`.

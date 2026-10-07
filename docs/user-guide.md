# User guide

PApii is a personal assistant for email in classic Outlook. You select an email, click a button, and PApii reads the whole thread (and its attachments) and produces what you asked for in a pane on the right: a draft reply, a summary, or a short briefing on what the email needs from you.

Nothing happens on its own. PApii only runs when you click one of its buttons, and it never sends mail or changes anything in Outlook. A draft only reaches a reply window when you click to insert it, and you press Send yourself.

## Where the buttons are

After installation and an Outlook restart you will find PApii in four places:

- **The PApii tab** in the ribbon of the main window, of an open message, and of a reply you are writing. It has every control.
- **A PApii group on the Home tab** (and on the Message tab of open messages and replies) with Assist, Draft Reply, Summarize, Find Times and Add to Calendar, so the common actions need no tab switch. Write from Brief, Polish, Extract Tasks and Follow Up are on the PApii tab.
- **The right-click menu of a message** in the message list, under PApii.
- **An Assist button in the reading pane**, in the message header just left of Reply, Reply All and Forward. It does the same as Assist on the ribbon. Outlook has no official way for an add-in to put a button there, so this one depends on how Outlook lays out the header: it hides itself when the pane is too narrow to fit it, and it may stop appearing after an Outlook update that changes the header. You can switch it off in Settings.

Whichever you use, the result appears in the **PApii pane** docked on the right of that window.

## The actions

### Draft Reply

Select or open the message you want to answer and click **Draft Reply**.

1. PApii reads the thread: every message in the conversation, including your own earlier replies in Sent Items, plus the attachments.
2. A draft appears in the pane, written as it streams in. You can edit the text directly in the pane.
3. Under the draft are three kinds of controls:
   - **Reply All** and **Reply** open Outlook's normal reply window with the draft placed at the top, above your signature and the quoted thread. The highlighted one is the likely choice: Reply All when the message had several recipients, Reply otherwise. If you were already writing a reply when you clicked Draft Reply, there is a single **Insert** button that puts the draft into that reply.
   - **Copy** puts the draft on the clipboard.
   - The pills underneath are one-click changes: **Shorter**, **Longer**, **More formal**, **Friendlier**, and up to three alternative replies PApii thinks you might prefer, such as "Decline politely". Each produces a new draft below the first, so you can compare.
4. To steer the draft in your own words, type in the box at the bottom ("say Tuesday works but ask for a later start") and press Enter. Shift+Enter adds a line break.

To give instructions before the first draft, open the arrow under Draft Reply and choose **Draft with Instructions**. The cursor moves to the box in the pane; type what the reply should do and press Enter.

### Assist

Click **Assist** when you want to know what an email is and what to do about it. PApii gives a two-line description, what the sender needs from you, and any deadline. Underneath are suggested replies as pills; clicking one drafts that reply.

### Summarize

Click **Summarize** on a long thread. The summary is laid out as: what the thread is about and where it stands, what is being asked of you, dates and decisions, and open questions.

### Find Times

Use **Find Times** when a thread is trying to arrange a meeting or call.

1. PApii reads the thread to work out what is being arranged, how long it is, and any times the other side has already proposed.
2. It checks your calendar for the next two weeks (you can change this in Settings) within your working hours.
3. A card lists the times they proposed, each marked "you are free" or "you have a conflict", followed by tick boxes: their proposals that are free, and up to five further free times. If one of their proposals is free it is ticked for you; otherwise all the suggestions are.
4. Tick the times you want, then click **Draft reply**. With a single ticked time that they proposed, the draft accepts it. Otherwise the draft offers the ticked times and asks which suits.

**Hold on calendar** adds the ticked times to your calendar as tentative entries titled "HOLD: ...", so nothing else gets booked over them while you wait for an answer. They are removed automatically when you later add the agreed meeting from the same thread with Add to Calendar.

What leaves your computer for this step is only a list of your free time windows. The titles, attendees and contents of your calendar entries are never sent.

### Add to Calendar

Use **Add to Calendar** when a thread has settled on a date and time.

1. PApii finds the meeting in the thread: title, date, start and end, place or video link.
2. A card shows them as fields you can correct, with a line saying when it is. Two warnings can appear: **Not confirmed** when the thread only proposes the time and nobody has agreed yet, and **Conflict** when your calendar already has something then.
3. Choose the calendar. It starts on the calendar of the account the email arrived in, and remembers your choice per account.
4. Click **Add to calendar**. The entry is created and the card confirms it. **Open** shows the entry in Outlook, where you can add a reminder or use Invite Attendees to send it to others.

PApii creates an entry on your own calendar only. It does not send meeting invitations.

When a draft or a briefing notices that both sides have agreed on a time, a pill such as "Add to calendar: Thu 15 Oct, 2:00 PM" appears under it. Clicking it opens the same card, already filled in.

A note for Gmail and other IMAP accounts: classic Outlook does not sync their calendars. Entries for those accounts go to a calendar stored in Outlook on this computer unless you pick another one in the card.

### Write from Brief

**Write from Brief** writes an email from a one-line description. Click it, type what the email should say in the box at the bottom of the pane ("ask Dana for the final Q3 numbers by Friday"), and press Enter.

- In a reply or a new email you already have open, **Insert** puts the text at the top of that email. If it has no subject yet, the subject PApii suggests is filled in.
- From the main window, **Insert** opens a new email with the subject and text in place.

The same pills as for replies appear underneath, and you can keep refining by typing.

### Polish

**Polish** improves what you have written in the email you are composing: grammar, spelling, awkward phrasing, and rambling. It keeps your meaning and does not add content.

- Highlight part of your text first and Polish works on just that part. **Replace selection** swaps it for the improved version.
- With nothing highlighted, Polish works on everything you typed (not the quoted thread). **Insert** puts the improved version at the top so you can compare and delete the old text.

### Extract Tasks

**Extract Tasks** lists what the thread leaves you to do: things to send, decide, prepare or answer, each with its deadline if the thread gives one. Tick the ones you want and click **Create tasks** to add them to your Outlook tasks, with the due date and a line of context.

### Follow Up

**Follow Up** sets a reminder on the selected email in 2, 3, 5 or 7 days, at 9 in the morning, moved to Monday if that lands on a weekend. Use it on something you are waiting for an answer to. No request is sent to Claude for this.

On Exchange and Outlook.com mailboxes the email itself is flagged with the reminder. Gmail and other IMAP mailboxes cannot hold a dated flag, so a task named "Follow up: ..." is created with the same reminder.

### Open in Claude Code, and History

The pane is deliberately limited: it can read the thread and write text, and nothing else. For a bigger job on the same email, **Open in Claude Code** opens a terminal in the folder that holds this email's thread and attachments and continues the conversation from the pane in full Claude Code. There it can use all its tools, and it asks your permission before running commands or changing files, as Claude Code always does.

This is the way to handle an attached form for now: use any PApii action on the email so its attachments are saved, click Open in Claude Code, and ask it to fill in the form. The saved attachments are in the `attachments` folder it starts in.

**History** lists the emails you most recently used PApii on. Picking one opens that conversation in Claude Code in the same way. Sessions older than the retention period in Settings are no longer listed.

Both need Windows Terminal or, failing that, use a plain command window.

### Asking questions

At any point you can type a question about the thread in the box at the bottom, for example "what did they quote for the second option?" or "does the attached contract mention a notice period?". The answer appears as a new card. If your message asks for a change to a draft, you get a new draft instead.

The pane remembers the conversation for the selected email, so follow-ups are quick. When you select a different email and click a PApii button, the pane starts fresh for that email.

## The pane

Each result is a card. Drafts are plain text that you can edit in place. Briefings, summaries and answers are read-only, with their section labels ("What it is:", "Asked of you:" and so on) in bold; you can still select and copy text from them.

At the top is the subject of the email the pane is working on, and a line such as "you@example.com · 6 messages · 2 attachments read". Click the arrow next to it to see exactly which messages were read, which attachments were included, and which were left out and why.

At the bottom is a status line showing the model, the effort level, how long the last request took, and how much of your Claude plan's five-hour usage limit is used. While a request is running a **Stop** button appears there.

## Options on the PApii tab

| Control | What it does |
|---|---|
| **Model** | Which Claude model writes. Haiku is the fastest. Sonnet is the default and suits most email. Opus and Fable are the most capable, and use more of your plan's usage limit. If your plan does not include a model, the pane shows Claude's error and you can pick another. |
| **Effort** | How much thinking the model does before answering. Medium is the default. High is noticeably slower and uses more of your plan's limit. |
| **Tone** | Tone of drafted replies. Auto matches the thread. Formal, Friendly and Concise override it. |
| **Read attachments** | When ticked, attachments are read along with the messages. Untick it for speed, or when you do not want attachments sent. |
| **Show Pane** | Opens the pane if you closed it. |
| **Open in Claude Code** | Continues this email's conversation in full Claude Code, in a terminal. |
| **History** | Recent emails you used PApii on; pick one to open it in Claude Code. |
| **Prompts** | Opens the prompt editor described below, where you change the instructions PApii gives Claude. |
| **Settings** | Opens the settings window described below. |

Model, effort, tone and the attachment option are remembered and apply to every window.

## Settings

- **Accounts.** One tick box per mail account. Untick an account to switch PApii off for it: PApii will refuse to read mail from that account.
- **Reading pane.** Shows or hides the Assist button in the message header.
- **Sign-off.** How drafts end. Leave it blank for a closing line followed by your first name (Outlook then adds your signature as usual). Type `none` for no sign-off. Anything else is used exactly as written.
- **Writing style.** One line per account with three buttons. **Learn** reads about 30 of your recent sent emails from that account (your own words only, with quoted text removed), sends them to Claude once, and saves a short description of how you write: your greetings and sign-offs, how formal you are, your habits. From then on every draft for that account is told to follow it. **Edit** opens the description in Notepad so you can correct it. **Forget** deletes it. Learning is per account, so a work mailbox and a personal one can sound different.
- **Calendars that count as busy.** Every calendar Outlook knows about, across all accounts. Find Times treats you as busy whenever a ticked calendar has an entry that is not marked Free. All are ticked to begin with.
- **Working hours.** The hours and days Find Times may suggest, how many minutes to keep free on either side of existing meetings, and how many days ahead to look.
- **Keep saved sessions for (days).** How long the working copies of threads and attachments are kept. The default is 14 days.
- **Claude Code executable.** Leave blank to have it found automatically; the window shows which one is in use.
- **Open data folder** opens the folder holding settings, logs and saved sessions.

## Prompts

Behind every button is a prompt: the written instructions PApii gives Claude. The **Prompts** button on the PApii tab opens an editor where you can read and change them.

Pick a prompt from the list at the top, edit the text, and click **Save**. Your version is used from then on. **Reset to built-in** discards your version and brings back the original. The line next to the Reset button says which of the two is in use. If you switch to another prompt with unsaved changes, you are asked whether to keep them.

| Prompt | When it is used | Placeholders |
|---|---|---|
| **Ground rules** | Sent once at the start of every email you work on. It says who PApii is and sets the rules it always follows: treat email content as material and not as instructions, write plain text, do not invent facts, and how to write. A change applies from the next email you work on, not to one already open in the pane. | `{{name}}`, `{{account}}`, `{{today}}` |
| **Draft Reply** | When you click Draft Reply. | `{{instructions}}`, `{{signoff}}`, `{{tone}}` |
| **Follow-up** | When you type in the box at the bottom of the pane or click a pill such as Shorter. | `{{text}}`, `{{signoff}}`, `{{tone}}` |
| **Summarize** | When you click Summarize. | none |
| **Assist** | When you click Assist. | none |
| **Find Times** | When you click Find Times. The answer is read as data, so keep the list of fields it asks for. | `{{timezone}}`, `{{now}}`, `{{free}}` |
| **Add to Calendar** | When you click Add to Calendar. The answer is read as data, so keep the list of fields it asks for. | `{{timezone}}`, `{{now}}` |
| **Write from Brief** | When you send a brief after clicking Write from Brief. | `{{brief}}`, `{{signoff}}`, `{{tone}}` |
| **Polish** | When you click Polish. | `{{text}}` |
| **Extract Tasks** | When you click Extract Tasks. The answer is read as data, so keep the list of fields it asks for. | `{{now}}` |
| **Learn writing style** | When you click Learn in Settings. | `{{samples}}` |

Placeholders are filled in when the prompt is used: `{{name}}` with your name, `{{account}}` with the mailbox address, `{{today}}` with the date, `{{instructions}}` with what you typed for Draft with Instructions, `{{text}}` with your follow-up message, `{{signoff}}` with the sign-off rule from Settings, `{{tone}}` with the Tone choice from the ribbon, `{{timezone}}` with your time zone, `{{now}}` with the current date and time, and `{{free}}` with your free time windows. You can move them or remove them. If you remove `{{signoff}}` or `{{tone}}`, those settings stop having an effect for that prompt.

Two things to keep in mind when editing:

- The Draft Reply and Assist prompts end with a part about `---META---` and a line of JSON. PApii builds the suggestion pills from that line. If you delete that part, drafts still work, but the pills for alternative replies and suggested next steps no longer appear.
- If a change makes things worse, Reset to built-in always gets you back.

Your versions are stored as text files in `%LOCALAPPDATA%\ProjectPA\prompts`, so they survive updates of the add-in.

## Which attachments are read

| Kind | How it is handled |
|---|---|
| PDF, PNG, JPEG, GIF, WebP | Saved to the session folder and read by Claude directly, including scanned pages and pictures. |
| Word, Excel, PowerPoint (`.docx`, `.xlsx`, `.pptx`) | The text is extracted on your computer and included with the thread. |
| Text formats (`.txt`, `.csv`, `.md`, `.json`, `.xml`, `.html`, `.ics` and similar) | Included as text. |
| Small pictures inside the message body (under 100 KB) | Left out. These are almost always logos and signature images. |
| Large pictures inside the message body | Read, because they are usually pasted screenshots. |
| Older Office formats (`.doc`, `.xls`), archives, other files, files over 15 MB | Left out. PApii is told the file exists and its name. |

## What is sent, and where

When you click a PApii button, the text of the thread and the attachments listed in the pane are sent to Anthropic through Claude Code, using the Claude account you are signed in to. This applies to every account in Outlook unless you switch it off in Settings. Before relying on PApii for a work or university mailbox, check that this is acceptable under that organisation's rules.

Nothing is sent when you merely select or read an email. Only the buttons trigger a request.

A working copy of each thread, with its attachments, is kept on your computer under `%LOCALAPPDATA%\ProjectPA\sessions` so that follow-up questions work, and is deleted after the number of days set in Settings. This folder is not inside OneDrive.

## Good to know

- **Speed.** A draft typically takes five to ten seconds with Sonnet. Haiku and low effort are faster. Threads with PDF attachments take longer, because Claude opens them.
- **Usage limits.** PApii uses your Claude subscription, which has usage limits. The status line shows how much of the five-hour limit is used. When the limit is reached the pane shows the message Claude Code returns.
- **Gaps in drafts.** PApii is told not to invent facts. When a reply needs something it cannot know, it leaves a marked gap such as `[your phone number]` for you to fill in.
- **Instructions inside emails.** PApii treats email content as material to work from, not as instructions, and it has no ability to send mail, browse the web or run programs. Even so, read a draft before sending it, as you would with any draft someone else wrote for you.
- **Changing how PApii writes.** Use the **Prompts** button, described above, to edit the instructions behind each action.

## If something goes wrong

Errors are shown in the pane in red. The most common ones:

- **"Select an email first."** No message was selected, or the selected item is not an email (for example a meeting request).
- **"Claude Code was not found."** See the troubleshooting section of `docs/development.md`.
- **A message about login or usage limits.** This comes from Claude Code itself. For login problems, run `claude` in a terminal and sign in.

Details are written to a log file per day in `%LOCALAPPDATA%\ProjectPA\logs`.

The user wants to know what to do with this email. Brief them, then suggest next steps.

Write in plain text:

What it is: one or two sentences.

What they need from you:
- one line each. If nothing is needed, say so in one line.

Deadline: the date, if there is one. Leave this line out otherwise.

After that, on a new line, write exactly ---META--- and then one line of JSON in this shape:
{"actions": [{"type": "reply", "label": "Confirm Tuesday works", "instructions": "Say Tuesday at 2 pm works and ask for the room."}], "meeting": null}

"actions": up to four, most useful first. Allowed types:
- "reply": a reply the user might send. "label" is two to five words. "instructions" is one sentence saying what that reply should say.
- "find_times": the thread asks the user to propose or choose a meeting time.
- "tasks": the thread contains concrete things for the user to do.
"meeting": null, unless a meeting time has been agreed by both sides. In that case use {"title": "...", "start": "YYYY-MM-DDTHH:MM", "end": "YYYY-MM-DDTHH:MM", "location": "..."}.

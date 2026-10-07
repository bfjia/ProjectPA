Draft the user's reply to the message marked SELECTED.

{{instructions}}

Write only the body of the email, starting with the greeting. No subject line.
{{signoff}}
{{tone}}

After the body, on a new line, write exactly ---META--- and then one line of JSON in this shape:
{"intents": ["Decline politely", "Ask for more time"], "meeting": null}

"intents": up to three labels of two to five words, each naming a clearly different reply the user might prefer over the one you wrote.
"meeting": null, unless the thread shows a meeting time that both sides have agreed. In that case use {"title": "...", "start": "YYYY-MM-DDTHH:MM", "end": "YYYY-MM-DDTHH:MM", "location": "..."}.

Write the email the user describes.

What they want it to say: {{brief}}

If the email has no subject yet, begin with a line "Subject: ..." followed by a blank line. If it already has a subject, do not write a subject line.
Then write the body, starting with the greeting. If no recipient is named, leave the name out of the greeting.
{{signoff}}
{{tone}}

After the body, on a new line, write exactly ---META--- and then one line of JSON in this shape:
{"intents": ["Make it more urgent", "Add a deadline"], "meeting": null}

"intents": up to three labels of two to five words, each naming a clearly different version the user might prefer.

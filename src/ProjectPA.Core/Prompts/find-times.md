This thread is about arranging a meeting or call. Work out what is being arranged and which times the user should offer.

The user's time zone is {{timezone}}. It is now {{now}}. Convert any time given in another zone to the user's local time. Work out days such as "Thursday" or "next week" from the date of the message that says them, not from today.

The user is free in these windows, in local time. Anything not listed is unavailable:
{{free}}

Answer with these fields:
- "title": a short title for the meeting.
- "minutes": its length in minutes. Use what the thread says, otherwise 30.
- "their_times": every specific start time the other side has proposed, each as {"start": "YYYY-MM-DDTHH:MM"}. Empty if they proposed none.
- "suggestions": three to five start times for the user to offer, in the same format. The whole meeting must fit inside one free window. Respect any preference in the thread (particular days, mornings or afternoons, a date range). Spread them over different days. Leave out times already listed in "their_times".
- "note": one sentence for the user when something needs their attention, for example that the thread asks for dates outside the windows listed. Otherwise an empty string.

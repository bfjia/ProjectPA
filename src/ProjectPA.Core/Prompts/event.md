Find the meeting, call or appointment this thread settles on, so the user can add it to their calendar.

The user's time zone is {{timezone}}. It is now {{now}}. Convert any time given in another zone to the user's local time. Work out days such as "Thursday" or "next week" from the date of the message that says them, not from today.

Answer with these fields:
- "found": true if the thread names a specific date and time for something the user should attend. If several were discussed, use the one finally agreed, or failing that the most recent proposal.
- "title": a short title.
- "start" and "end": local time as "YYYY-MM-DDTHH:MM". If no end or length is given, assume 30 minutes for a call and 60 for a meeting.
- "location": the room, address or video-call link. Empty if none is given.
- "notes": two or three lines worth having in the calendar entry: who it is with and what it is about.
- "evidence": the sentence in the thread that the date and time come from, copied word for word in quotation marks, then who wrote it and the date of their message in brackets. For example: "Does Thursday at 2 pm work for you?" (Dana Lee, 6 Oct). The user checks your answer against it.
- "confirmed": true if both sides have agreed to this time, false if it is only proposed.

If "found" is false, leave the other text fields empty.

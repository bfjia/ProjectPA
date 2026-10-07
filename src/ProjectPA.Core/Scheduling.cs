using System.Globalization;

namespace ProjectPA;

public class Busy
{
    public DateTime Start, End;
    public bool Hold;   // one of our own tentative holds
}

public class CalendarInfo
{
    public string Id { get; set; }     // store id | folder id
    public string Name { get; set; }
}

public class EventDraft
{
    public string Title = "", Location = "", Notes = "", CalendarId;
    public DateTime Start, End;   // wall-clock time in TimeZoneId, or local time when that is null
    public string TimeZoneId;
    public bool Hold;   // tentative placeholder for a time on offer
}

public static class Scheduling
{
    const string Times = """{"type":"array","items":{"type":"object","properties":{"start":{"type":"string"}},"required":["start"]}}""";

    public const string TimesSchema = """{"type":"object","properties":{"title":{"type":"string"},"minutes":{"type":"integer"},"their_times":""" + Times
        + ""","suggestions":""" + Times + ""","note":{"type":"string"}},"required":["title","minutes","their_times","suggestions","note"]}""";

    public const string EventSchema = """{"type":"object","properties":{"found":{"type":"boolean"},"title":{"type":"string"},"start":{"type":"string"},"end":{"type":"string"},"location":{"type":"string"},"notes":{"type":"string"},"confirmed":{"type":"boolean"}},"required":["found","title","start","end","location","notes","confirmed"]}""";

    public const string TasksSchema = """{"type":"object","properties":{"tasks":{"type":"array","items":{"type":"object","properties":{"title":{"type":"string"},"due":{"type":"string"},"notes":{"type":"string"}},"required":["title","due","notes"]}}},"required":["tasks"]}""";

    // the format the prompts ask for
    public static DateTime? Parse(string iso) =>
        DateTime.TryParse(iso, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    // A wall-clock time in another zone, as this computer's local time. Throws for a time the zone skips (clocks going forward).
    public static DateTime ToLocal(DateTime time, string zoneId) => zoneId == null ? time
        : TimeZoneInfo.ConvertTime(DateTime.SpecifyKind(time, DateTimeKind.Unspecified), TimeZoneInfo.FindSystemTimeZoneById(zoneId), TimeZoneInfo.Local);

    public static bool IsFree(IEnumerable<Busy> busy, DateTime start, DateTime end, int buffer = 0) =>
        !busy.Any(b => b.Start < end.AddMinutes(buffer) && b.End > start.AddMinutes(-buffer));

    // Free stretches inside working hours, day by day. Busy blocks are padded by the buffer; scraps under 15 minutes are dropped.
    public static List<Busy> FreeWindows(IEnumerable<Busy> busy, DateTime from, DateTime to, Settings s)
    {
        var free = new List<Busy>();
        var blocks = busy.Select(b => new Busy { Start = b.Start.AddMinutes(-s.BufferMinutes), End = b.End.AddMinutes(s.BufferMinutes) })
            .OrderBy(b => b.Start).ToList();
        for (var day = from.Date; day <= to.Date; day = day.AddDays(1))
        {
            if (!s.WorkDays.Contains(day.DayOfWeek.ToString().Substring(0, 3))) continue;
            var at = Later(day + TimeSpan.Parse(s.WorkStart), from);
            var end = Earlier(day + TimeSpan.Parse(s.WorkEnd), to);
            foreach (var b in blocks)
            {
                if (b.End <= at || b.Start >= end) continue;
                Add(at, b.Start);
                at = Later(at, b.End);
            }
            Add(at, end);
        }
        return free;

        void Add(DateTime a, DateTime b)
        {
            if ((b - a).TotalMinutes >= 15) free.Add(new Busy { Start = a, End = b });
        }
    }

    // one line per day, for the prompt
    public static string Describe(IEnumerable<Busy> windows) => string.Join("\n", windows.GroupBy(w => w.Start.Date)
        .Select(g => $"- {g.Key:ddd d MMM yyyy}: " + string.Join(", ", g.Select(w => $"{w.Start:HH:mm}-{w.End:HH:mm}"))));

    static DateTime Later(DateTime a, DateTime b) => a > b ? a : b;
    static DateTime Earlier(DateTime a, DateTime b) => a < b ? a : b;
}

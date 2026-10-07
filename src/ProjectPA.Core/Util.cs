using Newtonsoft.Json;

namespace ProjectPA;

public static class Paths
{
    // PROJECTPA_DATA: point tests and experiments somewhere else
    public static readonly string Data = Environment.GetEnvironmentVariable("PROJECTPA_DATA") ?? Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ProjectPA");
    public static string Logs => Dir("logs");
    public static string Sessions => Dir("sessions");
    static string Dir(string name) => Directory.CreateDirectory(Path.Combine(Data, name)).FullName;
}

public static class Log
{
    static readonly object gate = new();
    public static void Info(string msg) => Write("INF", msg);
    public static void Error(string msg, Exception e = null) => Write("ERR", e == null ? msg : $"{msg}: {e}");

    static void Write(string level, string msg)
    {
        try
        {
            lock (gate)
                File.AppendAllText(Path.Combine(Paths.Logs, $"{DateTime.Now:yyyy-MM-dd}.log"),
                    $"{DateTime.Now:HH:mm:ss.fff} {level} {msg}\r\n");
        }
        catch { }
    }
}

public class Settings
{
    public string Model = "sonnet", Effort = "medium", Tone = "auto", ClaudePath;
    public string SignOff = "";               // "": closing plus first name. "none": nothing. Else used as written.
    public bool IncludeAttachments = true;
    public bool HeaderButton = true;          // Assist button in the reading pane header
    public int KeepDays = 14;                 // session folders older than this are deleted
    public List<string> DisabledAccounts = new();

    // scheduling
    public string WorkStart = "09:00", WorkEnd = "17:00";
    public List<string> WorkDays = new() { "Mon", "Tue", "Wed", "Thu", "Fri" };
    public int BufferMinutes = 15;            // kept free either side of existing meetings
    public int HorizonDays = 14;              // how far ahead to look for free time
    public List<string> AvailabilityCalendars = new();          // calendar ids that count as busy; empty = all
    public Dictionary<string, string> EventCalendars = new();   // account -> calendar last chosen for its events

    static readonly string file = Path.Combine(Paths.Data, "settings.json");
    public static Settings Current = Load();

    static Settings Load()
    {
        try
        {
            // Replace: by default the saved list would be appended to the defaults above
            return JsonConvert.DeserializeObject<Settings>(File.ReadAllText(file),
                new JsonSerializerSettings { ObjectCreationHandling = ObjectCreationHandling.Replace }) ?? new();
        }
        catch { return new(); }
    }

    public void Save()
    {
        Directory.CreateDirectory(Paths.Data);
        File.WriteAllText(file, JsonConvert.SerializeObject(this, Formatting.Indented));
    }
}

public class PromptInfo
{
    public string Name { get; set; }
    public string Title { get; set; }
    public string About { get; set; }
}

public static class Prompts
{
    const string Meta = " Keep the ---META--- part: the suggestion pills are built from it.";

    // what the prompt editor lists
    public static readonly PromptInfo[] All =
    {
        new() { Name = "system", Title = "Ground rules", About = "Sent once at the start of every email you work on: who PApii is and the rules it always follows. Placeholders: {{name}}, {{account}}, {{today}}. A change applies from the next email." },
        new() { Name = "draft-reply", Title = "Draft Reply", About = "The request behind Draft Reply. Placeholders: {{instructions}}, {{signoff}}, {{tone}}." + Meta },
        new() { Name = "followup", Title = "Follow-up", About = "Used when you type in the pane or click a pill such as Shorter. Placeholders: {{text}}, {{signoff}}, {{tone}}." },
        new() { Name = "summarize", Title = "Summarize", About = "The request behind Summarize. No placeholders." },
        new() { Name = "assist", Title = "Assist", About = "The request behind Assist. No placeholders." + Meta },
        new() { Name = "find-times", Title = "Find Times", About = "The request behind Find Times. The answer is read as data, so keep the list of fields it asks for. Placeholders: {{timezone}}, {{now}}, {{free}} (your free windows)." },
        new() { Name = "event", Title = "Add to Calendar", About = "The request behind Add to Calendar. The answer is read as data, so keep the list of fields it asks for. Placeholders: {{timezone}}, {{now}}." },
    };

    static string CustomFile(string name) => Path.Combine(Paths.Data, "prompts", name + ".md");
    public static bool IsCustom(string name) => File.Exists(CustomFile(name));

    public static string BuiltIn(string name)
    {
        using var r = new StreamReader(typeof(Prompts).Assembly.GetManifestResourceStream(name + ".md"));
        return r.ReadToEnd().Trim();
    }

    // the user's version if there is one, else built-in; {{key}} gets replaced
    public static string Get(string name, params (string key, string value)[] fill)
    {
        var text = IsCustom(name) ? File.ReadAllText(CustomFile(name)).Trim() : BuiltIn(name);
        foreach (var (key, value) in fill) text = text.Replace("{{" + key + "}}", value ?? "");
        return text.Trim();
    }

    // empty or same as built-in: no custom file, so built-in updates keep arriving
    public static void Save(string name, string text)
    {
        text = (text ?? "").Replace("\r\n", "\n").Trim();
        if (text.Length == 0 || text == BuiltIn(name).Replace("\r\n", "\n")) { Reset(name); return; }
        Directory.CreateDirectory(Path.GetDirectoryName(CustomFile(name)));
        File.WriteAllText(CustomFile(name), text);
    }

    public static void Reset(string name)
    {
        if (IsCustom(name)) File.Delete(CustomFile(name));
    }

    public static string Tone(string tone) => tone switch
    {
        "formal" => "Tone: formal and polished.",
        "friendly" => "Tone: friendly and relaxed.",
        "concise" => "Tone: brief. Use as few words as politeness allows.",
        _ => "",
    };

    public static string SignOff(string signOff) => (signOff ?? "").Trim() switch
    {
        "" => "End with a short closing line, then the user's first name on the line directly below it. Add nothing after that: Outlook appends the signature.",
        "none" => "Do not add a closing line or a name.",
        var s => "End with exactly this sign-off:\n" + s,
    };
}

using Newtonsoft.Json;

namespace ProjectPA;

public static class Paths
{
    public static readonly string Data = Path.Combine(
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
    public int KeepDays = 14;                 // session folders older than this are deleted
    public List<string> DisabledAccounts = new();

    static readonly string file = Path.Combine(Paths.Data, "settings.json");
    public static Settings Current = Load();

    static Settings Load()
    {
        try { return JsonConvert.DeserializeObject<Settings>(File.ReadAllText(file)) ?? new(); }
        catch { return new(); }
    }

    public void Save()
    {
        Directory.CreateDirectory(Paths.Data);
        File.WriteAllText(file, JsonConvert.SerializeObject(this, Formatting.Indented));
    }
}

public static class Prompts
{
    // file in %LOCALAPPDATA%\ProjectPA\prompts beats built-in text; {{key}} gets replaced
    public static string Get(string name, params (string key, string value)[] fill)
    {
        var custom = Path.Combine(Paths.Data, "prompts", name + ".md");
        string text;
        if (File.Exists(custom)) text = File.ReadAllText(custom);
        else
        {
            using var r = new StreamReader(typeof(Prompts).Assembly.GetManifestResourceStream(name + ".md"));
            text = r.ReadToEnd();
        }
        foreach (var (key, value) in fill) text = text.Replace("{{" + key + "}}", value ?? "");
        return text.Trim();
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

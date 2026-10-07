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
    public string Model = "sonnet", Effort = "medium", ClaudePath;

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
    // file in %LOCALAPPDATA%\ProjectPA\prompts beats built-in text
    public static string Get(string name)
    {
        var custom = Path.Combine(Paths.Data, "prompts", name + ".md");
        if (File.Exists(custom)) return File.ReadAllText(custom);
        using var r = new StreamReader(typeof(Prompts).Assembly.GetManifestResourceStream(name + ".md"));
        return r.ReadToEnd();
    }
}

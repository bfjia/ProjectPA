using System.Diagnostics;
using System.Text;
using Newtonsoft.Json.Linq;

namespace ProjectPA;

public class ClaudeRequest
{
    public string Prompt, SystemPrompt, WorkDir, SessionId, JsonSchema;
    public string Model = "sonnet", Effort = "medium", Tools = "Read";
    public bool Resume;
}

public class ClaudeResult
{
    public string Text, SessionId, Error;
    public JToken Structured;
    public long Ms;
    public bool Ok => Error == null;
}

public class RateInfo
{
    public bool Limited;
    public double FiveHour, SevenDay;   // 0..1 of the window used
    public DateTime Resets;
}

// Runs the Claude Code CLI headless. Uses whatever login the CLI has.
public static class Claude
{
    public static string Find()
    {
        var set = Settings.Current.ClaudePath;
        if (File.Exists(set)) return set;

        foreach (var d in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';'))
            try { var f = Path.Combine(d.Trim(), "claude.exe"); if (File.Exists(f)) return f; } catch { }

        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var local = Path.Combine(home, ".local", "bin", "claude.exe");
        if (File.Exists(local)) return local;

        // VS Code extension ships one; folder name changes each update, take newest
        var ext = Path.Combine(home, ".vscode", "extensions");
        return !Directory.Exists(ext) ? null : Directory.GetDirectories(ext, "anthropic.claude-code-*")
            .Select(d => Path.Combine(d, "resources", "native-binary", "claude.exe"))
            .Where(File.Exists).OrderByDescending(File.GetLastWriteTimeUtc).FirstOrDefault();
    }

    public static string Args(ClaudeRequest r)
    {
        var a = new List<string>
        {
            "-p", "--model", r.Model, "--effort", r.Effort, "--tools", r.Tools,
            // no shell, web, MCP, skills or user settings: email text is untrusted
            "--permission-mode", "dontAsk", "--strict-mcp-config", "--restricted", "--disable-slash-commands",
        };
        a.AddRange(r.JsonSchema != null
            ? new[] { "--output-format", "json", "--json-schema", r.JsonSchema }
            : new[] { "--output-format", "stream-json", "--verbose", "--include-partial-messages" });
        if (r.SystemPrompt != null) a.AddRange(new[] { "--system-prompt-file", "system.md" });
        if (r.SessionId == null) a.Add("--no-session-persistence");
        else a.AddRange(new[] { r.Resume ? "--resume" : "--session-id", r.SessionId });
        return string.Join(" ", a.Select(Quote));
    }

    // onText: text delta, or null when a new assistant message starts (clear what was shown)
    public static async Task<ClaudeResult> Run(ClaudeRequest r, Action<string> onText = null,
        CancellationToken ct = default, Action<RateInfo> onRate = null)
    {
        var res = new ClaudeResult();
        var sw = Stopwatch.StartNew();
        try
        {
            var exe = Find();
            if (exe == null)
            {
                res.Error = "Claude Code was not found. Install it, or set its path in Settings.";
                return res;
            }
            Directory.CreateDirectory(r.WorkDir);
            if (r.SystemPrompt != null) File.WriteAllText(Path.Combine(r.WorkDir, "system.md"), r.SystemPrompt);

            var psi = new ProcessStartInfo(exe, Args(r))
            {
                WorkingDirectory = r.WorkDir, UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8,
            };
            // an API key in the environment would bill the API instead of the subscription
            psi.EnvironmentVariables.Remove("ANTHROPIC_API_KEY");
            psi.EnvironmentVariables.Remove("ANTHROPIC_AUTH_TOKEN");

            using var p = Process.Start(psi);
            using var stop = ct.Register(() => { try { p.Kill(); } catch { } });
            var err = p.StandardError.ReadToEndAsync();

            // raw bytes: the StandardInput writer would use the console code page
            var bytes = Encoding.UTF8.GetBytes(r.Prompt);
            await p.StandardInput.BaseStream.WriteAsync(bytes, 0, bytes.Length).ConfigureAwait(false);
            p.StandardInput.Close();

            string line;
            while ((line = await p.StandardOutput.ReadLineAsync().ConfigureAwait(false)) != null)
                Feed(line, res, onText, onRate);
            p.WaitForExit();

            if (ct.IsCancellationRequested) res.Error = "Stopped.";
            else if (res.Text == null && res.Error == null)
            {
                var e = (await err.ConfigureAwait(false)).Trim();
                res.Error = e.Length > 0 ? e : $"Claude exited with code {p.ExitCode}.";
            }
        }
        catch (Exception e)
        {
            res.Error = ct.IsCancellationRequested ? "Stopped." : e.Message;
            Log.Error("claude run", e);
        }
        finally { res.Ms = sw.ElapsedMilliseconds; }
        if (!res.Ok) Log.Error($"claude: {res.Error}");
        return res;
    }

    // one line of CLI output
    public static void Feed(string line, ClaudeResult res, Action<string> onText = null, Action<RateInfo> onRate = null)
    {
        if (line.Length == 0 || line[0] != '{') return;
        JObject o;
        try { o = JObject.Parse(line); } catch { return; }

        switch ((string)o["type"])
        {
            case "stream_event":
                var ev = o["event"];
                if ((string)ev?["type"] == "message_start") onText?.Invoke(null);
                else if ((string)ev?.SelectToken("delta.type") == "text_delta") onText?.Invoke((string)ev.SelectToken("delta.text"));
                break;

            case "rate_limit_event":
                var i = o["rate_limit_info"];
                onRate?.Invoke(new RateInfo
                {
                    Limited = (string)i?["status"] != "allowed",
                    FiveHour = (double?)i?.SelectToken("unifiedWindows.five_hour.utilization") ?? 0,
                    SevenDay = (double?)i?.SelectToken("unifiedWindows.seven_day.utilization") ?? 0,
                    Resets = DateTimeOffset.FromUnixTimeSeconds((long?)i?["resetsAt"] ?? 0).LocalDateTime,
                });
                break;

            case "result":
                res.SessionId = (string)o["session_id"];
                res.Structured = o["structured_output"];
                var text = (string)o["result"];
                if ((bool?)o["is_error"] == true) res.Error = string.IsNullOrEmpty(text) ? "Claude reported an error." : text;
                else res.Text = text ?? "";
                break;
        }
    }

    // Windows argv quoting
    public static string Quote(string s)
    {
        if (s.Length > 0 && s.IndexOfAny(new[] { ' ', '\t', '"' }) < 0) return s;
        var sb = new StringBuilder("\"");
        var slashes = 0;
        foreach (var c in s)
        {
            if (c == '\\') { slashes++; continue; }
            sb.Append('\\', c == '"' ? slashes * 2 + 1 : slashes).Append(c);
            slashes = 0;
        }
        return sb.Append('\\', slashes * 2).Append('"').ToString();
    }
}

namespace ProjectPA.Tests;

public class ClaudeTests
{
    // shapes recorded from claude 2.1.292, trimmed
    static readonly string[] Stream =
    {
        """{"type":"system","subtype":"init","session_id":"s1","tools":["Read"],"model":"claude-haiku-4-5-20251001"}""",
        """{"type":"stream_event","event":{"type":"message_start","message":{"role":"assistant","content":[]}},"session_id":"s1"}""",
        """{"type":"stream_event","event":{"type":"content_block_delta","index":0,"delta":{"type":"thinking_delta","thinking":""}},"session_id":"s1"}""",
        """{"type":"rate_limit_event","rate_limit_info":{"status":"allowed","resetsAt":1791400800,"rateLimitType":"five_hour","unifiedWindows":{"five_hour":{"utilization":0.21,"resetsAt":1791400800},"seven_day":{"utilization":0.03,"resetsAt":1791925200}}},"session_id":"s1"}""",
        """{"type":"stream_event","event":{"type":"content_block_delta","index":1,"delta":{"type":"text_delta","text":"Hello from"}},"session_id":"s1"}""",
        """{"type":"stream_event","event":{"type":"content_block_delta","index":1,"delta":{"type":"text_delta","text":" the probe."}},"session_id":"s1"}""",
        """{"type":"assistant","message":{"role":"assistant","content":[{"type":"text","text":"Hello from the probe."}]},"session_id":"s1"}""",
        """{"type":"result","subtype":"success","is_error":false,"result":"Hello from the probe.","session_id":"s1","duration_ms":817}""",
    };

    [Fact]
    public void Feed_streams_text_then_result()
    {
        var res = new ClaudeResult();
        var shown = "x";
        RateInfo rate = null;
        foreach (var line in Stream)
            Claude.Feed(line, res, t => shown = t == null ? "" : shown + t, r => rate = r);

        Assert.Equal("Hello from the probe.", shown);   // message_start cleared the "x"
        Assert.Equal("Hello from the probe.", res.Text);
        Assert.Equal("s1", res.SessionId);
        Assert.True(res.Ok);
        Assert.False(rate.Limited);
        Assert.Equal(0.21, rate.FiveHour);
    }

    [Fact]
    public void Feed_reads_structured_output()
    {
        var res = new ClaudeResult();
        Claude.Feed("""{"type":"result","is_error":false,"result":"{}","structured_output":{"title":"Budget","minutes":30},"session_id":"s2"}""", res);
        Assert.Equal(30, (int)res.Structured["minutes"]);
    }

    [Fact]
    public void Feed_reports_errors_and_ignores_noise()
    {
        var res = new ClaudeResult();
        Claude.Feed("", res);
        Claude.Feed("not json", res);
        Claude.Feed("{broken", res);
        Assert.Null(res.Text);
        Claude.Feed("""{"type":"result","is_error":true,"result":"Usage limit reached"}""", res);
        Assert.False(res.Ok);
        Assert.Equal("Usage limit reached", res.Error);
    }

    [Theory]
    [InlineData("plain", "plain")]
    [InlineData("", "\"\"")]
    [InlineData("two words", "\"two words\"")]
    [InlineData("{\"a\":1}", "\"{\\\"a\\\":1}\"")]
    [InlineData("C:\\dir with space\\", "\"C:\\dir with space\\\\\"")]
    public void Quote_follows_windows_rules(string raw, string quoted) => Assert.Equal(quoted, Claude.Quote(raw));

    [Fact]
    public void Sessions_are_listed_and_removed()
    {
        var dir = Directory.CreateDirectory(Path.Combine(Paths.Sessions, "20261007-114000-123")).FullName;
        File.WriteAllText(Path.Combine(dir, "thread.md"), "# Email thread: Q3 <budget> & more\nMailbox: x\n");
        var bare = Directory.CreateDirectory(Path.Combine(Paths.Sessions, "20261007-120000-000")).FullName;   // no thread.md: not listed

        var s = Context.RecentSessions().Single(x => x.dir == dir);
        Assert.Equal("Q3 <budget> & more", s.subject);
        Assert.Equal(new DateTime(2026, 10, 7, 11, 40, 0), s.when);
        Assert.DoesNotContain(Context.RecentSessions(), x => x.dir == bare);

        // forget: false keeps the test from starting Claude Code
        Assert.Equal(2, Context.RemoveSessions(new[] { dir, bare }, forget: false));
        Assert.False(Directory.Exists(dir));
        Assert.False(Directory.Exists(bare));
    }

    [Fact]
    public void Settings_defaults_and_reload_from_disk()
    {
        var s = new Settings();
        Assert.Equal(("opus", "medium", 7), (s.Model, s.Effort, s.KeepDays));

        var file = Path.Combine(Paths.Data, "settings.json");
        Settings.Current.Tone = "formal";
        Settings.Current.Save();
        Assert.Equal("formal", Settings.Current.Tone);

        // someone else edits the file: the next read sees it, and the lists are replaced, not appended to
        File.WriteAllText(file, File.ReadAllText(file).Replace("\"formal\"", "\"concise\""));
        File.SetLastWriteTimeUtc(file, DateTime.UtcNow.AddSeconds(5));
        Assert.Equal("concise", Settings.Current.Tone);
        Assert.Equal(5, Settings.Current.WorkDays.Count);

        Settings.Current.Tone = "auto";
        Settings.Current.Save();
    }

    [Fact]
    public void Unreadable_settings_are_flagged_and_not_overwritten()
    {
        var file = Path.Combine(Directory.CreateDirectory(Paths.Data).FullName, "settings.json");
        try
        {
            File.WriteAllText(file, "{ \"Model\": ");
            File.SetLastWriteTimeUtc(file, DateTime.UtcNow.AddSeconds(10));
            Assert.NotNull(Settings.Current.Broken);        // not defaults in silence: those would switch every account back on

            Settings.Current.Save();                        // as a ribbon dropdown would
            Assert.Equal("{ \"Model\": ", File.ReadAllText(file));

            var s = Settings.Current;
            s.Broken = null;                                // as saving from the Settings window does
            s.Save();
            Assert.Null(Settings.Current.Broken);
            Assert.Contains("\"Model\": \"opus\"", File.ReadAllText(file));
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public void Args_pick_mode_and_session()
    {
        var stream = Claude.Args(new ClaudeRequest { SessionId = "id1", SystemPrompt = "x" });
        Assert.Contains("--output-format stream-json", stream);
        Assert.Contains("--session-id id1", stream);
        Assert.Contains("--system-prompt-file system.md", stream);
        Assert.Contains("--restricted", stream);

        var json = Claude.Args(new ClaudeRequest { JsonSchema = "{\"type\":\"object\"}" });
        Assert.Contains("--output-format json", json);
        Assert.Contains("--no-session-persistence", json);
        Assert.DoesNotContain("--system-prompt-file", json);

        Assert.Contains("--resume id1", Claude.Args(new ClaudeRequest { SessionId = "id1", Resume = true }));
    }
}

using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;

namespace ProjectPA;

// Turns a thread into the text Claude gets, and picks apart what comes back.
public static class Context
{
    public const int MaxBody = 20_000, MaxThread = 150_000, MaxAttachText = 30_000;
    public const long MaxAttachBytes = 15 * 1024 * 1024;
    public const int MinInlinePicture = 100 * 1024;
    public const string MetaMark = "---META---";

    static readonly string[] ClaudeReads = { ".pdf", ".png", ".jpg", ".jpeg", ".gif", ".webp" };
    static readonly string[] PlainText = { ".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".htm", ".html", ".ics", ".vcf", ".log", ".yml", ".yaml" };
    const RegexOptions I = RegexOptions.IgnoreCase;
    static readonly Regex Wrote = new(@"(wrote|a écrit|schrieb|escribió|ha scritto|schreef)\s*:$", I);

    public static string NewSession() =>
        Directory.CreateDirectory(Path.Combine(Paths.Sessions, $"{DateTime.Now:yyyyMMdd-HHmmss-fff}")).FullName;

    public static void PurgeSessions(int days)
    {
        foreach (var d in Directory.GetDirectories(Paths.Sessions))
            try { if (Directory.GetCreationTime(d) < DateTime.Now.AddDays(-days)) Directory.Delete(d, true); }
            catch (Exception e) { Log.Error("purge " + d, e); }
    }

    // Cuts quoted history off a reply. Only for messages whose predecessors are in the thread.
    public static string StripQuotes(string body)
    {
        var lines = (body ?? "").Replace("\r\n", "\n").Split('\n');
        var end = lines.Length;
        for (var i = 0; i < lines.Length && end == lines.Length; i++)
        {
            var l = lines[i].Trim();
            var next = i + 1 < lines.Length ? lines[i + 1].Trim() : "";
            var header = string.Join(" ", lines.Skip(i + 1).Take(4));   // Outlook's From/Sent/To block
            if (Regex.IsMatch(l, @"^-{2,}\s*Original Message\s*-{2,}$", I)
                || Regex.IsMatch(l, @"^(From|De|Von|Da|Van):\s", I) && Regex.IsMatch(header, @"\b(Sent|Date|Enviado|Gesendet|Envoyé|Inviato|Verzonden|Datum):\s", I)
                // "On <date> <name> wrote:", which Gmail wraps onto a second line
                || l.Length < 300 && Regex.IsMatch(l, @"^(On|Le|Am|El|Il|Op)\s", I) && (Wrote.IsMatch(l) || Wrote.IsMatch(next))
                // ">" block, unless answers are interleaved with it
                || l.StartsWith(">") && lines.Skip(i).All(x => x.Trim().Length == 0 || x.TrimStart().StartsWith(">")))
                end = i;
        }
        // blank lines and Outlook's rule of underscores above the header
        while (end > 0 && lines[end - 1].Trim().All(c => c == '_')) end--;
        return string.Join("\n", lines.Take(end)).TrimEnd();
    }

    // After the host saved attachments: pull text out of what we can, keep what Claude reads itself, drop the rest.
    public static void Digest(EmailThread t, string workDir)
    {
        foreach (var a in t.Messages.SelectMany(m => m.Attachments).Where(a => a.SavedAs != null))
        {
            var file = Path.Combine(workDir, a.SavedAs);
            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (ClaudeReads.Contains(ext)) continue;
            a.Text = ExtractText(file);
            if (a.Text == null) a.Note = "file type not read";
            else if (a.Text.Length > MaxAttachText) a.Text = a.Text.Substring(0, MaxAttachText) + "\n[cut: attachment text too long]";
            a.SavedAs = null;
            try { File.Delete(file); } catch { }
        }
    }

    public static string ExtractText(string file)
    {
        try
        {
            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (PlainText.Contains(ext)) return File.ReadAllText(file);
            if (ext is not (".docx" or ".pptx" or ".xlsx")) return null;
            using var zip = ZipFile.OpenRead(file);   // Office files are zipped XML
            return ext == ".docx" ? Paragraphs(zip, n => n == "word/document.xml")
                : ext == ".pptx" ? Paragraphs(zip, n => Regex.IsMatch(n, @"^ppt/slides/slide\d+\.xml$"))
                : Sheets(zip);
        }
        catch (Exception e) { Log.Error("extract " + file, e); }
        return null;
    }

    static XDocument Xml(ZipArchiveEntry e) { using var s = e.Open(); return XDocument.Load(s); }
    static IEnumerable<XElement> Named(this XContainer x, string name) => x.Descendants().Where(e => e.Name.LocalName == name);

    // Word and PowerPoint both use <p> paragraphs holding <t> text runs
    static string Paragraphs(ZipArchive zip, Func<string, bool> pick) => string.Join("\n\n",
        zip.Entries.Where(e => pick(e.FullName)).OrderBy(e => e.FullName.Length).ThenBy(e => e.FullName)
            .Select(e => string.Join("\n", Xml(e).Named("p").Select(p => string.Concat(p.Named("t").Select(t => t.Value))))));

    static string Sheets(ZipArchive zip)
    {
        var shared = zip.GetEntry("xl/sharedStrings.xml") is { } ss
            ? Xml(ss).Root.Elements().Select(si => string.Concat(si.Named("t").Select(t => t.Value))).ToList()
            : new List<string>();
        var sb = new StringBuilder();
        foreach (var e in zip.Entries.Where(e => Regex.IsMatch(e.FullName, @"^xl/worksheets/sheet\d+\.xml$")).OrderBy(e => e.FullName.Length).ThenBy(e => e.FullName))
        {
            sb.AppendLine($"[{Path.GetFileNameWithoutExtension(e.Name)}]");
            foreach (var row in Xml(e).Named("row").Take(500))
                sb.AppendLine(string.Join("\t", row.Elements().Select(c =>
                {
                    var v = c.Elements().FirstOrDefault(x => x.Name.LocalName == "v")?.Value;
                    return (string)c.Attribute("t") == "s" && int.TryParse(v, out var i) && i < shared.Count
                        ? shared[i] : v ?? string.Concat(c.Named("t").Select(x => x.Value));
                })));
        }
        return sb.ToString();
    }

    public static string Render(EmailThread t)
    {
        var parts = t.Messages.Select((m, i) => RenderMessage(t, m, i)).ToList();
        var skip = 0;   // too long: oldest go first
        while (skip < parts.Count - 1 && parts.Skip(skip).Sum(p => p.Length) > MaxThread) skip++;

        var sb = new StringBuilder();
        sb.AppendLine($"# Email thread: {t.Subject}");
        sb.AppendLine($"Mailbox: {t.Account}. The user is {t.UserName}.");
        sb.AppendLine($"{t.Messages.Count} message(s), oldest first.");
        if (skip > 0) sb.AppendLine($"[{skip} oldest message(s) left out to save space]");
        foreach (var p in parts.Skip(skip)) sb.AppendLine().Append(p);
        return sb.ToString();
    }

    static string RenderMessage(EmailThread t, Message m, int i)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"## Message {i + 1} of {t.Messages.Count}{(m.Selected ? " (SELECTED)" : "")}");
        sb.AppendLine($"From: {(m.Mine ? "the user" : m.From)}");
        if (!string.IsNullOrEmpty(m.To)) sb.AppendLine($"To: {m.To}");
        if (!string.IsNullOrEmpty(m.Cc)) sb.AppendLine($"Cc: {m.Cc}");
        sb.AppendLine($"Date: {m.Sent:ddd yyyy-MM-dd HH:mm}");
        if (Bare(m.Subject) != Bare(t.Subject)) sb.AppendLine($"Subject: {m.Subject}");
        if (m.Attachments.Count > 0)
        {
            sb.AppendLine("Attachments:");
            foreach (var a in m.Attachments)
                sb.AppendLine($"- {a.Name} ({a.Size / 1024 + 1} KB): " + (
                    a.SavedAs != null ? $"read it with the Read tool at {a.SavedAs}"
                    : a.Text != null ? "text included below"
                    : $"not included ({a.Note})"));
        }

        // a forward carries content the thread may not have; the first message has nothing before it
        var body = i == 0 || Regex.IsMatch(m.Subject ?? "", @"^\s*(fw|fwd):", I) ? (m.Body ?? "").Trim() : StripQuotes(m.Body);
        if (body.Length > MaxBody) body = body.Substring(0, MaxBody) + "\n[cut: message too long]";
        sb.AppendLine().AppendLine(body);
        foreach (var a in m.Attachments.Where(a => a.Text != null))
            sb.AppendLine().AppendLine($"### Attachment text: {a.Name}").AppendLine(a.Text.Trim());
        return sb.ToString();
    }

    static string Bare(string subject) => Regex.Replace(subject ?? "", @"^\s*((re|fw|fwd|aw|sv|tr):\s*)+", "", I).Trim();

    // Splits "body, MetaMark, json". meta is null when absent or unparseable.
    public static (string body, JObject meta) SplitMeta(string text)
    {
        var i = text.IndexOf(MetaMark, StringComparison.Ordinal);
        if (i < 0) return (text.Trim(), null);
        JObject meta = null;
        var json = text.Substring(i + MetaMark.Length);
        int a = json.IndexOf('{'), b = json.LastIndexOf('}');
        if (a >= 0 && b > a) try { meta = JObject.Parse(json.Substring(a, b - a + 1)); } catch { }
        return (text.Substring(0, i).Trim(), meta);
    }

    // What to show mid-stream: no marker, not even a half-arrived one.
    public static string Visible(string partial)
    {
        var i = partial.IndexOf(MetaMark, StringComparison.Ordinal);
        if (i >= 0) return partial.Substring(0, i).TrimEnd();
        for (var n = Math.Min(MetaMark.Length - 1, partial.Length); n > 0; n--)
            if (partial.EndsWith(MetaMark.Substring(0, n), StringComparison.Ordinal))
                return partial.Substring(0, partial.Length - n);
        return partial;
    }
}

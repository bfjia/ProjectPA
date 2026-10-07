using System.IO.Compression;

namespace ProjectPA.Tests;

public class ContextTests
{
    [Theory]
    // Outlook header block under a rule
    [InlineData("Sounds good.\n\n________________________________\nFrom: Dana Lee <dana@x.com>\nSent: Monday, October 5, 2026 2:03 PM\nTo: Me\nSubject: RE: Budget\n\nOld text", "Sounds good.")]
    // Gmail, wrapped
    [InlineData("Yes, Tuesday works.\n\nOn Mon, Oct 5, 2026 at 3:14 PM Dana Lee <dana@x.com>\nwrote:\n\n> Old text\n> more", "Yes, Tuesday works.")]
    [InlineData("Thanks.\n\n-----Original Message-----\nFrom: someone", "Thanks.")]
    [InlineData("Merci.\n\nLe lun. 5 oct. 2026 à 15:14, Dana <dana@x.com> a écrit :\n> ancien", "Merci.")]
    // trailing ">" block
    [InlineData("See below.\n\n> quoted one\n> quoted two", "See below.")]
    public void StripQuotes_cuts_history(string body, string kept) => Assert.Equal(kept, Context.StripQuotes(body));

    [Theory]
    [InlineData("No quotes here.\nSecond line.")]
    // answers between quoted lines must survive
    [InlineData("> Can you do Tuesday?\nYes.\n> And the budget?\nAttached.")]
    // "From:" in running text, no header block after it
    [InlineData("From: the top of my head, I think so.\nLet me check.")]
    [InlineData("On Monday I will send it.\nThanks")]
    public void StripQuotes_leaves_the_rest(string body) => Assert.Equal(body, Context.StripQuotes(body));

    static EmailThread Sample() => new()
    {
        Subject = "Budget", Account = "me@x.com", UserName = "Sam Jones",
        Messages =
        {
            new Message { From = "Dana Lee <dana@x.com>", To = "Sam Jones", Subject = "Budget", Sent = new DateTime(2026, 10, 5, 9, 0, 0),
                Body = "Can you review?\n\nOn Fri someone\nwrote:\n> kept, first message is never stripped",
                Attachments =
                {
                    new Attachment { Name = "plan.pdf", Size = 2048, SavedAs = "attachments/1-plan.pdf" },
                    new Attachment { Name = "notes.docx", Size = 100, Text = "Line one" },
                    new Attachment { Name = "logo.png", Size = 10, Note = "inline image" },
                } },
            new Message { Mine = true, Subject = "RE: Budget", Sent = new DateTime(2026, 10, 5, 10, 0, 0),
                Body = "Will do.\n\nFrom: Dana Lee\nSent: Monday\n\nCan you review?" },
            new Message { From = "Dana Lee <dana@x.com>", Subject = "FW: Budget (revised)", Sent = new DateTime(2026, 10, 6, 8, 0, 0), Selected = true,
                Body = "FYI\n\nFrom: Finance\nSent: Tuesday\n\nForwarded content stays" },
        },
    };

    [Fact]
    public void Render_describes_the_thread()
    {
        var text = Context.Render(Sample());
        Assert.Contains("# Email thread: Budget", text);
        Assert.Contains("The user is Sam Jones", text);
        Assert.Contains("## Message 3 of 3 (SELECTED)", text);
        Assert.Contains("From: the user", text);
        Assert.Contains("plan.pdf (3 KB): read it with the Read tool at attachments/1-plan.pdf", text);
        Assert.Contains("notes.docx (1 KB): text included below", text);
        Assert.Contains("### Attachment text: notes.docx\nLine one".Replace("\n", Environment.NewLine), text);
        Assert.Contains("not included (inline image)", text);
        Assert.Contains("first message is never stripped", text);
        Assert.Contains("Will do.", text);
        Assert.DoesNotContain("Sent: Monday", text);             // reply: quote cut
        Assert.Contains("Forwarded content stays", text);        // forward: kept
        Assert.Contains("Subject: FW: Budget (revised)", text);  // differs from the thread subject
        Assert.DoesNotContain("Subject: RE: Budget", text);      // RE: alone is not a change
    }

    [Fact]
    public void Render_drops_oldest_when_too_long()
    {
        var t = Sample();
        t.Messages[0].Body = new string('a', Context.MaxBody);
        t.Messages[1].Body = new string('b', Context.MaxBody);
        for (var i = 0; i < 10; i++)
            t.Messages.Insert(0, new Message { From = "x", Subject = "Budget", Body = new string('c', Context.MaxBody) });
        var text = Context.Render(t);
        Assert.True(text.Length <= Context.MaxThread + 1000);
        Assert.Contains("oldest message(s) left out", text);
        Assert.Contains("(SELECTED)", text);
    }

    [Fact]
    public void SplitMeta_separates_body_and_json()
    {
        var (body, meta) = Context.SplitMeta("Hi Dana,\n\nYes.\n\nSam\n---META---\n{\"intents\":[\"Decline\"],\"meeting\":null}\n");
        Assert.Equal("Hi Dana,\n\nYes.\n\nSam", body);
        Assert.Equal("Decline", (string)meta["intents"][0]);

        Assert.Null(Context.SplitMeta("Just an answer.").meta);
        var broken = Context.SplitMeta("Body\n---META---\n{oops");
        Assert.Equal("Body", broken.body);
        Assert.Null(broken.meta);
    }

    [Theory]
    [InlineData("Hello", "Hello")]
    [InlineData("Hello\n---META---\n{\"a\"", "Hello")]
    [InlineData("Hello\n---ME", "Hello\n")]      // marker half arrived
    [InlineData("Hello\n-", "Hello\n")]
    [InlineData("a - b", "a - b")]
    public void Visible_hides_the_marker(string partial, string shown) => Assert.Equal(shown, Context.Visible(partial));

    static readonly DateTime Mon = new(2026, 10, 12);   // a Monday

    static Busy Block(int day, double from, double to) => new() { Start = Mon.AddDays(day).AddHours(from), End = Mon.AddDays(day).AddHours(to) };

    [Fact]
    public void FreeWindows_respect_hours_days_and_buffers()
    {
        var s = new Settings { WorkStart = "09:00", WorkEnd = "17:00", BufferMinutes = 15 };
        var busy = new[] { Block(0, 10, 11), Block(0, 11.25, 12), Block(1, 8, 9.5), Block(1, 16.5, 18), Block(2, 0, 24) };

        // Monday 08:00 to Sunday: Saturday and Sunday are not working days
        var text = Scheduling.Describe(Scheduling.FreeWindows(busy, Mon.AddHours(8), Mon.AddDays(6), s));
        var lines = text.Split('\n');
        Assert.Equal("- Mon 12 Oct 2026: 09:00-09:45, 12:15-17:00", lines[0]);   // the 15-minute gap between meetings vanishes into the buffers
        Assert.Equal("- Tue 13 Oct 2026: 09:45-16:15", lines[1]);
        Assert.Equal("- Thu 15 Oct 2026: 09:00-17:00", lines[2]);                 // Wednesday is blocked all day
        Assert.Equal("- Fri 16 Oct 2026: 09:00-17:00", lines[3]);
        Assert.Equal(4, lines.Length);

        // starting mid-morning cuts the first window; no buffer leaves the gap
        s.BufferMinutes = 0;
        var later = Scheduling.FreeWindows(busy, Mon.AddHours(9.5), Mon.AddHours(23), s);
        Assert.Equal("- Mon 12 Oct 2026: 09:30-10:00, 11:00-11:15, 12:00-17:00", Scheduling.Describe(later));
    }

    [Fact]
    public void IsFree_checks_overlap_with_buffer()
    {
        var busy = new[] { Block(0, 10, 11) };
        Assert.True(Scheduling.IsFree(busy, Mon.AddHours(11), Mon.AddHours(12)));
        Assert.False(Scheduling.IsFree(busy, Mon.AddHours(11), Mon.AddHours(12), buffer: 15));
        Assert.False(Scheduling.IsFree(busy, Mon.AddHours(10.5), Mon.AddHours(11.5)));
        Assert.True(Scheduling.IsFree(busy, Mon.AddHours(8), Mon.AddHours(9.5), buffer: 15));
    }

    [Fact]
    public void Parse_reads_the_prompt_format() =>
        Assert.Equal(new DateTime(2026, 10, 15, 14, 0, 0), Scheduling.Parse("2026-10-15T14:00"));

    [Fact]
    public void Schemas_are_valid_json()
    {
        Assert.Equal("object", (string)Newtonsoft.Json.Linq.JObject.Parse(Scheduling.TimesSchema)["type"]);
        Assert.Equal("object", (string)Newtonsoft.Json.Linq.JObject.Parse(Scheduling.EventSchema)["type"]);
    }

    [Theory]
    [InlineData("Subject: Q3 numbers\n\nHi Dana,\n\nCould you send them?", "Q3 numbers", "Hi Dana,\n\nCould you send them?")]
    [InlineData("subject:  Hello \r\n\r\nBody", "Hello", "Body")]
    [InlineData("Hi Dana,\n\nSubject: this is not a header", "", "Hi Dana,\n\nSubject: this is not a header")]
    [InlineData("Just text", "", "Just text")]
    public void SplitSubject_peels_a_leading_subject_line(string text, string subject, string body) =>
        Assert.Equal((subject, body), Context.SplitSubject(text));

    [Fact]
    public void StyleNote_is_empty_until_learned_and_file_names_are_safe()
    {
        Assert.Equal("", Prompts.StyleNote("nobody@example.com"));
        Assert.EndsWith(@"style\a_b@example.com.md", Prompts.StyleFile("a/b@example.com"));
        var file = Prompts.StyleFile("writer@example.com");
        Directory.CreateDirectory(Path.GetDirectoryName(file));
        File.WriteAllText(file, "- Opens with Hi\n");
        Assert.EndsWith("Follow this in anything you draft for them:\n- Opens with Hi", Prompts.StyleNote("writer@example.com"));
    }

    static string Marked(string line) => string.Concat(Context.Runs(line).Select(r => r.bold ? $"<{r.text}>" : r.text));

    [Theory]
    [InlineData("What it is: Dana wants a decision.", "<What it is:> Dana wants a decision.")]
    [InlineData("What they need from you:", "<What they need from you:>")]
    [InlineData("Deadline: **20 October**, firm", "<Deadline:> <20 October>, firm")]
    [InlineData("Résumé: en bref", "<Résumé:> en bref")]
    [InlineData("- Confirm that Thursday 2:30 works", "- Confirm that Thursday 2:30 works")]   // list item, and a time
    [InlineData("See https://example.com/a for details", "See https://example.com/a for details")]
    [InlineData("the total is 10,100: more than planned", "the total is 10,100: more than planned")]   // not at line start
    [InlineData("", "")]
    public void Runs_bold_labels_and_marked_text(string line, string marked) => Assert.Equal(marked, Marked(line));

    [Fact]
    public void ExtractText_reads_office_files()
    {
        var dir = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "pa-" + Guid.NewGuid())).FullName;
        try
        {
            string Make(string name, params (string entry, string xml)[] parts)
            {
                var file = Path.Combine(dir, name);
                using var zip = ZipFile.Open(file, ZipArchiveMode.Create);
                foreach (var (entry, xml) in parts)
                    using (var w = new StreamWriter(zip.CreateEntry(entry).Open())) w.Write(xml);
                return file;
            }

            var docx = Make("a.docx", ("word/document.xml",
                "<w:document xmlns:w='w'><w:body><w:p><w:r><w:t>Dear </w:t></w:r><w:r><w:t>Sam</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>"));
            Assert.Equal("Dear Sam\nSecond", Context.ExtractText(docx));

            var pptx = Make("a.pptx",
                ("ppt/slides/slide10.xml", "<p:sld xmlns:p='p' xmlns:a='a'><a:p><a:r><a:t>Ten</a:t></a:r></a:p></p:sld>"),
                ("ppt/slides/slide2.xml", "<p:sld xmlns:p='p' xmlns:a='a'><a:p><a:r><a:t>Two</a:t></a:r></a:p></p:sld>"));
            Assert.Equal("Two\n\nTen", Context.ExtractText(pptx));

            var xlsx = Make("a.xlsx",
                ("xl/sharedStrings.xml", "<sst xmlns='s'><si><t>Item</t></si><si><t>Cost</t></si></sst>"),
                ("xl/worksheets/sheet1.xml", "<worksheet xmlns='s'><sheetData><row><c t='s'><v>0</v></c><c t='s'><v>1</v></c></row><row><c t='inlineStr'><is><t>Desk</t></is></c><c><v>120</v></c></row></sheetData></worksheet>"));
            Assert.Equal("[sheet1]\r\nItem\tCost\r\nDesk\t120\r\n", Context.ExtractText(xlsx));

            File.WriteAllText(Path.Combine(dir, "a.csv"), "x,y");
            Assert.Equal("x,y", Context.ExtractText(Path.Combine(dir, "a.csv")));
            File.WriteAllText(Path.Combine(dir, "a.bin"), "??");
            Assert.Null(Context.ExtractText(Path.Combine(dir, "a.bin")));
        }
        finally { Directory.Delete(dir, true); }
    }
}

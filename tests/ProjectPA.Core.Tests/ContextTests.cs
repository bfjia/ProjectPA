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

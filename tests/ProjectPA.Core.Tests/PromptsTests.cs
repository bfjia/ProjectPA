using System.Runtime.CompilerServices;

// the data folder and the settings in it are shared by every test
[assembly: CollectionBehavior(DisableTestParallelization = true)]

// net48 lacks the attribute; the compiler only needs the name
namespace System.Runtime.CompilerServices
{
    [AttributeUsage(AttributeTargets.Method)]
    sealed class ModuleInitializerAttribute : Attribute { }
}

namespace ProjectPA.Tests
{
    static class TestData
    {
        // before anything touches Paths: keep tests out of the real data folder
        [ModuleInitializer]
        internal static void Init() => Environment.SetEnvironmentVariable("PROJECTPA_DATA",
            Path.Combine(Path.GetTempPath(), "ProjectPA-tests-" + Guid.NewGuid()));
    }

    public class PromptsTests
    {
        [Fact]
        public void Data_folder_is_redirected() => Assert.Contains("ProjectPA-tests-", Paths.Data);

        [Fact]
        public void Every_listed_prompt_has_built_in_text()
        {
            foreach (var p in Prompts.All) Assert.False(string.IsNullOrWhiteSpace(Prompts.BuiltIn(p.Name)), p.Name);
        }

        [Fact]
        public void Get_fills_placeholders()
        {
            var text = Prompts.Get("followup", ("text", "make it shorter"), ("signoff", "S"), ("tone", "T"));
            Assert.Contains("The user says: make it shorter", text);
            Assert.DoesNotContain("{{", text);
        }

        [Fact]
        public void Save_then_reset_round_trips()
        {
            const string name = "summarize";
            var builtIn = Prompts.BuiltIn(name);
            try
            {
                Prompts.Save(name, "My own summary prompt.\r\n");
                Assert.True(Prompts.IsCustom(name));
                Assert.Equal("My own summary prompt.", Prompts.Get(name));

                // saving the built-in text, or nothing, means "no custom version"
                Prompts.Save(name, builtIn);
                Assert.False(Prompts.IsCustom(name));
                Prompts.Save(name, "x");
                Prompts.Save(name, "   ");
                Assert.False(Prompts.IsCustom(name));

                Prompts.Save(name, "again");
                Prompts.Reset(name);
                Assert.Equal(builtIn, Prompts.Get(name));
            }
            finally { Prompts.Reset(name); }
        }

        [Fact]
        public void Reset_without_a_custom_file_is_fine() => Prompts.Reset("assist");
    }
}

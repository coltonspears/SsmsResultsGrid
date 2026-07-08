using System.IO;
using System.Threading;
using SsmsResultsGrid.Core.Export;
using Xunit;

namespace SsmsResultsGrid.Core.Tests
{
    public class CsvWriterTests
    {
        private static string WriteToString(string[] columns, Models.ResultRow[] rows)
        {
            using var writer = new StringWriter();
            CsvWriter.Write(writer, columns, rows, CancellationToken.None);
            return writer.ToString();
        }

        [Fact]
        public void PlainValues_WrittenUnquoted()
        {
            var text = WriteToString(new[] { "a", "b" }, Rows.Make(new[] { "1", "2" }));
            Assert.Equal("a,b" + System.Environment.NewLine + "1,2" + System.Environment.NewLine, text);
        }

        [Fact]
        public void CommasQuotesAndNewlines_AreQuotedAndEscaped()
        {
            var text = WriteToString(
                new[] { "col" },
                Rows.Make(new[] { "va\"l,ue" }, new[] { "line1\nline2" }));

            Assert.Contains("\"va\"\"l,ue\"", text);
            Assert.Contains("\"line1\nline2\"", text);
        }

        [Fact]
        public void RaggedRow_PadsMissingCellsAsEmpty()
        {
            var text = WriteToString(new[] { "a", "b", "c" }, Rows.Make(new[] { "only" }));
            Assert.Contains("only,,", text);
        }

        [Fact]
        public void Escape_LeavesSafeValuesAlone()
        {
            Assert.Equal("plain", CsvWriter.Escape("plain"));
            Assert.Equal(string.Empty, CsvWriter.Escape(null));
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using SsmsResultsGrid.Core.Models;

namespace SsmsResultsGrid.Core.Export
{
    /// <summary>RFC-4180 style CSV export of the currently visible rows.</summary>
    public static class CsvWriter
    {
        public static void Write(
            TextWriter writer,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<ResultRow> rows,
            CancellationToken ct)
        {
            if (writer == null) throw new ArgumentNullException(nameof(writer));
            if (columnNames == null) throw new ArgumentNullException(nameof(columnNames));
            if (rows == null) throw new ArgumentNullException(nameof(rows));

            var line = new StringBuilder();
            for (int c = 0; c < columnNames.Count; c++)
            {
                if (c > 0) line.Append(',');
                line.Append(Escape(columnNames[c]));
            }
            writer.WriteLine(line.ToString());

            for (int r = 0; r < rows.Count; r++)
            {
                if ((r & 0x0FFF) == 0) ct.ThrowIfCancellationRequested();
                line.Clear();
                for (int c = 0; c < columnNames.Count; c++)
                {
                    if (c > 0) line.Append(',');
                    line.Append(Escape(rows[r][c]));
                }
                writer.WriteLine(line.ToString());
            }
        }

        public static void WriteFile(
            string path,
            IReadOnlyList<string> columnNames,
            IReadOnlyList<ResultRow> rows,
            CancellationToken ct)
        {
            using (var stream = new StreamWriter(path, append: false, encoding: new UTF8Encoding(encoderShouldEmitUTF8Identifier: true)))
            {
                Write(stream, columnNames, rows, ct);
            }
        }

        internal static string Escape(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            bool needsQuoting = value.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0;
            if (!needsQuoting) return value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace shtxt
{
    public class TextWriter
    {
        public static void Write(string name, IEnumerable<IReadOnlyCollection<string>> outputs, Config config)
        {
            var ext = "";
            var separator = "";
            switch (config.TextFormatType)
            {
                case TextFormatType.Csv:
                    ext = ".csv";
                    separator = ",";
                    break;
                case TextFormatType.Tsv:
                    ext = ".tsv";
                    separator = "\t";
                    break;
                default:
                    throw new Exception("unimplemented text format");
            }

            using (var writer = new StreamWriter(Path.Combine(config.OutputDir, name + ext)))
            {
                switch (config.NewLineType)
                {
                    case NewLineType.CR:
                        writer.NewLine = "\r";
                        break;
                    case NewLineType.LF:
                        writer.NewLine = "\n";
                        break;
                    case NewLineType.CRLF:
                        writer.NewLine = "\r\n";
                        break;
                }

                var lines = outputs.Select(data => String.Join(separator, data));
                foreach (var line in lines)
                {
                    writer.WriteLine(line);
                }
            }
        }

    }
}
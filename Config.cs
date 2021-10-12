// unset

using System.Collections.Generic;
using System.IO;
using System.Text;

using YamlDotNet.RepresentationModel;
using YamlDotNet.Serialization;

namespace Shtxt
{
    public enum NewLine
    {
        LF,
        CR,
        CRLF,
    }

    public enum TextFormat
    {
        Csv,
        Tsv,
        Yaml,
        Json,
    }
    public class Config
    {
        public string OutputDir { get; set; } = ".";
        public NewLine NewLine { get; set; } = NewLine.LF;

        public void SetNewLine(string str)
        {
            NewLine l = NewLine.LF;
            NewLine.TryParse(str, true, out l);
            this.NewLine = l;
        }
        public TextFormat TextFormat { get; set; } = TextFormat.Tsv;

        public void SetTextFormat(string str)
        {
            TextFormat t;
            TextFormat.TryParse(str, true, out t);
            this.TextFormat = t;
        }
        public string CommentStartsWith { get; set; } = "#";
        public string TableNameTag { get; set; } = "[テーブル名]";
        public string ColumnControlTag { get; set; } = "[カラム制御]";
        public string ColumnNameTag { get; set; } = "[カラム名]";

        public static Config LoadFromFile(string path)
        {
            using (var reader = new StreamReader(path, Encoding.UTF8))
            {
                var deserializer = new DeserializerBuilder().Build();
                return deserializer.Deserialize<Config>(reader.ReadToEnd());
            }

        }
    }
}
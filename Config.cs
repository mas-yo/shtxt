// unset

using System.Collections.Generic;
using System.IO;
using System.Text;

using YamlDotNet.RepresentationModel;

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
        public TextFormat TextFormat { get; set; } = TextFormat.Tsv;
        public string CommentStartsWith { get; set; } = "#";
        public string TableNameTag { get; set; } = "[テーブル名]";
        public string ColumnControlTag { get; set; } = "[カラム制御]";
        public string ColumnNameTag { get; set; } = "[カラム名]";

        public void LoadFromFile(string path)
        {
            using (var reader = new StreamReader(path, Encoding.UTF8))
            {
                var yaml = new YamlStream();
                yaml.Load(reader);
                foreach (var node in yaml.Documents[0].AllNodes)
                {
                    if (node.NodeType == YamlNodeType.Mapping)
                    {
                        var mapping = node as YamlMappingNode;
                        foreach (var child in mapping)
                        {
                            switch (child.Key.ToString())
                            {
                                case "OutputDir":
                                    OutputDir = child.Value.ToString();
                                break;
                                case "NewLine":
                                    NewLine l = NewLine.LF;
                                    NewLine.TryParse(child.Value.ToString(), true, out l);
                                    this.NewLine = l;
                                    break;
                                case "TextFormat":
                                    TextFormat t;
                                    TextFormat.TryParse(child.Value.ToString(), true, out t);
                                    this.TextFormat = t;
                                    break;
                                case "CommentStartsWith":
                                    CommentStartsWith = child.Value.ToString();
                                    break;
                                case "TableNameTag":
                                    TableNameTag = child.Value.ToString();
                                    break;
                                case "ColumnControlTag":
                                    ColumnControlTag = child.Value.ToString();
                                    break;
                                case "ColumnNameTag":
                                    ColumnNameTag = child.Value.ToString();
                                    break;
                            }
                        }
                    }
                }
            }

        }
    }
}
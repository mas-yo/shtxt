using System;
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
        public List<DirectoryInfo> InputFiles { get; set; }
        public string OutputDir { get; set; } = ".";
        public string NewLine { get; set; } = "LF";

        // public void SetNewLine(string str)
        // {
        //     NewLine l = NewLine.LF;
        //     NewLine.TryParse(str, true, out l);
        //     this.NewLine = l;
        // }
        public string TextFormat { get; set; } = "tsv";

        // public void SetTextFormat(string str)
        // {
        //     TextFormat t;
        //     TextFormat.TryParse(str, true, out t);
        //     this.TextFormat = t;
        // }
        public string CommentStartsWith { get; set; } = "#";
        public string TableNameTag { get; set; } = "[テーブル名]";
        public string ColumnControlTag { get; set; } = "[カラム制御]";
        public string ColumnNameTag { get; set; } = "[カラム名]";
        
        public FileInfo ConfigFile { get; set; }

        private void SetByKeyValue(string key, string value)
        {
            if (key == "InputFiles")
            {
                if (InputFiles == null) InputFiles = new List<DirectoryInfo>();
                
                var files = value.Split(" ");
                foreach (var file in files)
                {
                    InputFiles.Add( new DirectoryInfo(file));
                }

                return;
            }
            var propertyInfo = this.GetType().GetProperty(key);
            if (propertyInfo == null) return;
            propertyInfo.SetValue(this, value);
        }
        public void LoadFromFile()
        {
            var configFileList = new List<string>() {"config.yml","config.yaml", "Config.yml", "Config.yaml"};
            string path = null;
            if (ConfigFile != null && ConfigFile.Exists)
            {
                path = ConfigFile.FullName;
            }
            else
            {
                foreach (var file in configFileList)
                {
                    if (File.Exists(file))
                    {
                        path = file;
                        break;
                    }
                }
            }

            if (String.IsNullOrEmpty(path)) return;
            
            using (var reader = new StreamReader(path, Encoding.UTF8))
            {
                var deserializer = new DeserializerBuilder().Build();
                var dict = deserializer.Deserialize<Dictionary<string, string>>(reader.ReadToEnd());
                foreach (var kv in dict)
                {
                    SetByKeyValue(kv.Key, kv.Value);
                }
            }
        }

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
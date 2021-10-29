using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using YamlDotNet.RepresentationModel;
using YamlDotNet.Serialization;

namespace shtxt
{
    public enum NewLineType
    {
        LF,
        CR,
        CRLF,
    }

    public enum TextFormatType
    {
        Csv,
        Tsv,
        Yaml,
        Json,
    }
    public class Config
    {
        public List<DirectoryInfo> InputFiles { get; set; }
        public string InputPattern { get; set; } = "";

        public FileInfo VersionList { get; set; } = new FileInfo("versions.txt");
        
        public string CurrentVersion { get; set; }
        public string OutputDir { get; set; } = ".";
        public string NewLine { get; set; } = "LF";

        public NewLineType NewLineType
        {
            get => Enum.Parse<NewLineType>(NewLine, true);
        }

        public string TextFormat { get; set; } = "tsv";

        public TextFormatType TextFormatType
        {
            get => Enum.Parse<TextFormatType>(TextFormat, true);
        }

        public string CommentStartsWith { get; set; } = "#";
        public string TableNameTag { get; set; } = "[NAME]";
        public string ColumnControlTag { get; set; } = "[CONTROL]";
        public string ColumnNameTag { get; set; } = "[COLUMN]";
        
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

            if (key == "VersionList" && !String.IsNullOrEmpty(value))
            {
                VersionList = new FileInfo(value);
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
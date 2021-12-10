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
        public List<FileSystemInfo> InputFiles { get; set; }
        public string InputPattern { get; set; } = "";

        public string ExcludeInputPattern { get; set; } = "";
        public FileInfo VersionList { get; set; } = new FileInfo("versions.txt");
        
        public string CurrentVersion { get; set; }
        public string OutputDir { get; set; } = ".";

        public string DateTimeFormat { get; set; } = "";
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

        public string OutputControlColumn { get; set; } = "false";

        public bool IsOutputControlColumn
        {
            get => bool.Parse(OutputControlColumn);
        }
        public string OutputColumnNameTag { get; set; } = "";
        public string OutputCommentTag { get; set; } = "";

        public string CommentStartsWith { get; set; } = "#";
        public string TableNameTag { get; set; } = "[NAME]";
        public string ColumnControlTag { get; set; } = "[CONTROL]";
        public string ColumnNameTag { get; set; } = "[COLUMN]";
        
        public FileInfo ConfigFile { get; set; }

        private void SetByKeyValue(string key, string value)
        {
            if (key == "InputFiles")
            {
                if (InputFiles == null) InputFiles = new List<FileSystemInfo>();
                
                var files = value.Split(" ");
                foreach (var file in files)
                {
                    FileSystemInfo dirInfo = new DirectoryInfo(file);
                    if (dirInfo.Exists)
                    {
                        InputFiles.Add(dirInfo);
                    }

                    FileSystemInfo fileInfo = new FileInfo(file);
                    if (fileInfo.Exists)
                    {
                        InputFiles.Add(fileInfo);
                    }
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
            var paths = new List<string>();

            foreach (var file in configFileList)
            {
                if (File.Exists(file))
                {
                    paths.Add(file);
                    break;
                }
            }
            
            if (ConfigFile != null && ConfigFile.Exists)
            {
                paths.Add(ConfigFile.FullName);
            }

            foreach (var path in paths)
            {
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
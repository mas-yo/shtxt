using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.CommandLine.Parsing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading.Tasks.Dataflow;

using ICSharpCode.SharpZipLib.Core;

using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Cms;

namespace shtxt
{
    class Program
    {
        static bool IsEnable(Control control, string currentVersion, IList<string> versionList)
        {
            switch (control)
            {
                case None:
                    return true;
                case Comment:
                    return false;
                    
                case Version(var comp, var version):
                    var currentIndex = versionList.IndexOf(currentVersion);
                    if (currentIndex < 0) return true;
                    var checkIndex = versionList.IndexOf(version);
                    switch (comp)
                    {
                        case Compairator.Less:
                            return currentIndex < checkIndex;
                        case Compairator.LessOrEqual:
                            return currentIndex <= checkIndex;
                        case Compairator.Greater:
                            return currentIndex > checkIndex;
                        case Compairator.GreaterOrEqual:
                            return currentIndex >= checkIndex;
                        case Compairator.Equal:
                            return currentIndex == checkIndex;
                        case Compairator.NotEqual:
                            return currentIndex != checkIndex;
                    }

                    return true;
            }
            return true;
        }
        static void WriteCsv(StreamWriter writer, SheetInfo info, string separator, IList<string> versionList, string currentVersion)
        {
            var lines = GetRowEnumerable(info, versionList, currentVersion).Select(data => String.Join(separator, data));
            foreach (var line in lines)
            {
                writer.WriteLine(line);
            }
        }
        
        static IEnumerable<IReadOnlyCollection<string>> GetRowEnumerable(SheetInfo info, IList<string> versionList, string currentVersion)
        {
            var columnInfos = info.GetEnumerableColumnInfo();
            var skipColumns = columnInfos
                .Select((columnInfo, idx) => (idx, IsEnable(columnInfo.Control, currentVersion, versionList)))
                .Where(i => i.Item2 == false)
                .Select(i => i.idx)
                .ToList();

            yield return columnInfos
                .Where((ci, idx) => !skipColumns.Contains(idx))
                .Select(ci => ci.Name)
                .ToList().AsReadOnly();

            foreach (var row in info.Body)
            {
                if (!IsEnable(row.Control, currentVersion, versionList)) continue;
                
                yield return row.Data
                    .Where((data, idx) => !skipColumns.Contains(idx))
                    .ToList()
                    .AsReadOnly();
            }
        }

        static void WriteText(SheetInfo info, Config config)
        {
            if (!info.IsValid)
                return;

            var versionList = new List<string>();
            if (File.Exists(config.VersionList.FullName))
            {
                versionList = File.ReadLines(config.VersionList.FullName).ToList();
            }

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

            using (var writer = new StreamWriter(Path.Combine(config.OutputDir, info.Header.Name + ext)))
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

                WriteCsv(writer, info, separator, versionList, config.CurrentVersion);
            }
        }
        
        static void Convert(Config config)
        {
            var regex = new Regex(config.InputPattern);
            
            var tasks = config.InputFiles.SelectMany(info =>
            {
                if ((info.Attributes & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    return Directory.GetFiles(info.FullName, "*", SearchOption.AllDirectories).Where(path => regex.IsMatch(path));
                }
                else
                {
                    if (regex.IsMatch(info.FullName))
                    {
                        return new string[] {info.FullName};
                    }
                    return new string[] {};
                }
            }).SelectMany(path =>
            {
                var book = WorkbookFactory.Create(path, password: null, readOnly: true);
                return book.GetSheetEnumerable();
            }).Select(sheet =>
            {
                return Task.Factory.StartNew(() =>
                {
                    var controlParser = new ControlParser() { CommentStartsWith = config.CommentStartsWith };
                    var loader = new SheetLoader(config.TableNameTag, config.ColumnControlTag, config.ColumnNameTag, controlParser);
                    var info = loader.Load(sheet.GetRowDataEnumerable());
                    WriteText(info, config);
                });
            });

            Task.WaitAll(tasks.ToArray());
        }

        static void Main(string[] args)
        {
            var command = new RootCommand();
            command.Add(new Argument<List<DirectoryInfo>>("input-files"));
            command.Add(new Option<string>(new string[] {"-p", "--input-pattern"}, "input file pattern"));
            command.Add(new Option<FileInfo>(new string[] {"-l", "--version-list"}, "output version list"));
            command.Add(new Option<string>(new string[] {"-r", "--current-version"}, "output current version"));
            command.Add(new Option<string>(new string[] {"-o", "--output-dir"}, "output directory"));
            command.Add(new Option<string>(new string[] {"-n", "--newline"}, "newline code(cr,lf,crlf)"));
            command.Add(new Option<string>(new string[] {"-f", "--text-format"}, "output format(csv,tsv,json,yaml)"));
            command.Add(new Option<string>(new string[] {"--comment-starts-with"}, "comment line letter"));
            command.Add(new Option<string>(new string[] {"--table-name-tag"}, "table name tag"));
            command.Add(new Option<string>(new string[] {"--column-name-tag"}, "column name tag"));
            command.Add(new Option<string>(new string[] {"--column-control-tag"}, "column control tag"));
            command.Add(new Option<FileInfo>(new string[] {"-c", "--config-file"}, "config file"));

            command.Handler = CommandHandler.Create((Config config) =>
            {
                config.LoadFromFile();
                Convert(config);
            });
            
            command.Invoke(args);
            Console.WriteLine("End");
        }
    }
}
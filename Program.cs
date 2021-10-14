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

using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace shtxt
{
    class Program
    {
        static void WriteCsv(StreamWriter writer, SheetInfo info, string separator, Config config)
        {
            writer.WriteLine(String.Join(separator, info.Header.ColumnNames));
            foreach (var row in info.Body)
            {
                if (row.Control != null && row.Control.StartsWith(config.CommentStartsWith)) continue;
                writer.WriteLine(String.Join(separator, row.Data));
            }
        }

        static void WriteText(SheetInfo info, Config config)
        {
            if (!info.IsValid)
                return;

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

                WriteCsv(writer, info, separator, config);
            }
        }
        
        static void Convert(Config config)
        {
            var regex = new Regex(config.InputFilePattern);
            
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
                    var loader = new SheetLoader(config.TableNameTag, config.ColumnControlTag, config.ColumnNameTag);
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
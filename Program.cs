using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace MasterDataConverter
{
    class Program
    {
        static void WriteCsv(StreamWriter writer, MasterTableInfo info, string separator, Config config)
        {
            writer.WriteLine(String.Join(separator, info.Header.ColumnNames));
            foreach (var row in info.Body)
            {
                if (row.Control != null && row.Control.StartsWith(config.CommentStartsWith)) continue;
                writer.WriteLine(String.Join(separator, row.Data));
            }
        }

        static void WriteText(MasterTableInfo info, Config config)
        {
            if (!info.IsValid)
                return;

            var ext = "";
            var separator = "";
            switch (config.TextFormat)
            {
                case TextFormat.Csv:
                    ext = ".csv";
                    separator = ",";
                    break;
                case TextFormat.Tsv:
                    ext = ".tsv";
                    separator = "\t";
                    break;
                default:
                    throw new Exception("unimplemented text format");
            }
            
            using (var writer = new StreamWriter(Path.Combine(config.OutputDir, info.Header.Name + ext)))
            {
                switch (config.NewLine)
                {
                    case NewLine.CR:
                        writer.NewLine = "\r";
                        break;
                    case NewLine.LF:
                        writer.NewLine = "\n";
                        break;
                    case NewLine.CRLF:
                        writer.NewLine = "\r\n";
                        break;
                }

                WriteCsv(writer, info, separator, config);
            }
        }

        static void Convert(IList<string> inputPaths, Config config)
        {
            var tasks = inputPaths.SelectMany(path =>
            {
                var book = WorkbookFactory.Create(path, password: null, readOnly: true);
                return book.GetSheetEnumerable();
            }).Select(sheet =>
            {
                return Task.Factory.StartNew(() =>
                {
                    var loader = new MasterTableLoader(config.TableNameTag, config.ColumnControlTag, config.ColumnNameTag);
                    var info = loader.Load(sheet.GetRowDataEnumerable());
                    WriteText(info, config);
                });
            });

            Task.WaitAll(tasks.ToArray());
        }

        static void Main(string[] args)
        {
            var command = new RootCommand();
            command.Add(new Argument<List<FileInfo>>("input-files"));
            command.Add(new Option<string>(new string[]{"-o", "--output-dir"}, "output directory"));
            command.Add(new Option<string>(new string[]{"-n", "--newline"}, "newline code(cr,lf,crlf)"));
            command.Add(new Option<string>(new string[]{"-f", "--format"}, "output format(csv,tsv,json,yaml)"));
            command.Add(new Option<string>(new string[]{"--comment-starts-with"}, "comment line letter"));
            command.Add(new Option<string>(new string[]{"--table-name-tag"}, "table name tag"));
            command.Add(new Option<string>(new string[]{"--column-name-tag"}, "column name tag"));
            command.Add(new Option<string>(new string[]{"--column-control-tag"}, "column control tag"));
            command.Add(new Option<string>(new string[]{"-c", "--config"}, "config file"));
            

            command.Handler = CommandHandler.Create((
                List<FileInfo> inputFiles,
                string outputDir,
                string newline,
                string format,
                string commentStartsWith,
                string tableNameTag,
                string columnNameTag,
                string columnControlTag,
                string config) =>
            {
                var cfg = new Config();
                
                // TODO overwrite configs
                cfg.OutputDir = "../../../Output";
                Convert(inputFiles.Select(f => f.FullName).ToList(), cfg);
            });

            command.Invoke(args);
            Console.WriteLine("End");
        }
    }

}
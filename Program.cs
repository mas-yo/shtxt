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
        static void WriteTsv(StreamWriter writer, MasterTableInfo info)
        {
            writer.WriteLine(String.Join("\t", info.ColumnNames));
            foreach (var row in info.Body)
            {
                writer.WriteLine(String.Join("\t", row.Data));
            }
        }

        static void Convert(IList<string> inputPaths, string outputDir)
        {
            var tasks = inputPaths.SelectMany(path =>
            {
                var book = WorkbookFactory.Create(path, password: null, readOnly: true);
                return book.GetSheetEnumerable();
            }).Select(sheet =>
            {
                return Task.Factory.StartNew(() =>
                {
                    var info = MasterTableInfo.Load(sheet.GetRowDataEnumerable());
                    if (info.IsValid)
                    {
                        using (var writer = new StreamWriter(Path.Combine(outputDir, info.Name + ".tsv")))
                        {
                            writer.NewLine = "\n";
                            WriteTsv(writer, info);
                        }
                    }
                });
            });

            Task.WaitAll(tasks.ToArray());
        }

        static void Main(string[] args)
        {
            var command = new RootCommand();
            command.Add(new Argument<List<FileInfo>>("input_files"));
            command.Add(new Option<string>("-o", "output directory"));

            command.Handler = CommandHandler.Create<List<FileInfo>, string>((input_files, o) =>
            {
                Convert(input_files.Select(f => f.FullName).ToList(), o);
            });

            command.Invoke(args);
            Console.WriteLine("End");
        }
    }

}
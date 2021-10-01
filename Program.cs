using System;
using System.Collections.Generic;
using System.IO;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

using NPOI.XSSF.UserModel;

namespace MasterDataConverter
{
    public static class POIExtensions
    {
        public static IEnumerable<ISheet> GetSheetEnumerable(this IWorkbook book)
        {
            var enm = book.GetEnumerator();
            while (enm.MoveNext())
            {
                yield return enm.Current;
            }
        }
        public static IEnumerable<IList<String>> GetRowDataEnumerable(this ISheet sheet)
        {
            var enm = sheet.GetRowEnumerator();
            while (enm.MoveNext())
            {
                var row = (IRow)enm.Current;
                var data = new List<String>();
                foreach (var cell in row.Cells)
                {
                    while (data.Count < cell.ColumnIndex)
                    {
                        data.Add("");
                    }

                    data.Add(cell.ToString());
                }

                yield return data;
            }
        }
    }

    class TextFileInfo
    {
        public bool IsValid = false;
        public string FileName = null;
        public string Header = null;
        public IList<string> DataLines = new List<string>();
    }

    class Program
    {
        static TextFileInfo ConvertToTextFileInfo(IEnumerable<IList<String>> rows)
        {
            var info = new TextFileInfo();
            
            foreach (var columns in rows)
            {
                if (columns.Count <= 0) continue;

                var command = columns[0];
                if (command.StartsWith("#")) continue;
                
                if (command == "[テーブル名]" && columns.Count >= 2)
                {
                    info.FileName = columns[1];
                }
                else if (command == "[カラム名]")
                {
                    info.Header = String.Join("\t", columns);
                }
                else
                {
                    if (info.Header != null)
                    {
                        info.DataLines.Add(String.Join("\t", columns));
                    }
                }
            }

            if (!String.IsNullOrEmpty(info.FileName) && !string.IsNullOrEmpty(info.Header) && info.DataLines.Count > 0)
            {
                info.IsValid = true;
            }

            return info;
        }

        static void Convert(IList<string> inputPaths, string outputDir)
        {
            var tasks = inputPaths.SelectMany(path =>
            {
                var book = WorkbookFactory.Create(path, password:null, readOnly:true);
                return book.GetSheetEnumerable();
            }).Select(sheet =>
            {
                return Task.Factory.StartNew(() =>
                {
                    var info = ConvertToTextFileInfo(sheet.GetRowDataEnumerable());
                    if (info.IsValid)
                    {
                        using (var writer = new StreamWriter(Path.Combine(outputDir, info.FileName)))
                        {
                            writer.WriteLine(info.Header);
                            foreach (var line in info.DataLines)
                            {
                                writer.WriteLine(line);
                            }
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

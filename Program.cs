using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using NPOI.SS.UserModel;

namespace shtxt
{
    class Program
    {
        static void Main(string[] args)
        {
            var command = new RootCommand();
            command.Add(new Argument<List<DirectoryInfo>>("input-files"));
            command.Add(new Option<string>(new string[] {"-p", "--input-pattern"}, "Input file name pattern"));
            command.Add(new Option<FileInfo>(new string[] {"-l", "--version-list"}, "Version list file"));
            command.Add(new Option<string>(new string[] {"-r", "--current-version"}, "Current version"));
            command.Add(new Option<string>(new string[] {"-o", "--output-dir"}, "Output directory"));
            command.Add(new Option<string>(new string[] {"-n", "--newline"}, "Newline code(cr,lf,crlf)"));
            command.Add(new Option<string>(new string[] {"-f", "--text-format"}, "Output format(csv,tsv)"));
            command.Add(new Option<string>(new string[] {"--comment-starts-with"}, "String that indicates comment"));
            command.Add(new Option<string>(new string[] {"--table-name-tag"}, "Tag string for table name"));
            command.Add(new Option<string>(new string[] {"--column-name-tag"}, "Tag string for column name"));
            command.Add(new Option<string>(new string[] {"--column-control-tag"}, "Tag string for column control"));
            command.Add(new Option<FileInfo>(new string[] {"-c", "--config-file"}, "Config file"));

            command.Handler = CommandHandler.Create((Config config) =>
            {
                config.LoadFromFile();
                Convert(config);
            });
            
            command.Invoke(args);
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
                    (var name, var outputs) = Converter.Convert(sheet, config);
                    if (String.IsNullOrEmpty(name)) return;
                    TextWriter.Write(name, outputs, config);
                });
            });

            Task.WaitAll(tasks.ToArray());
        }
        
    }
}
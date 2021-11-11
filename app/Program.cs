using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
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
            command.Add(new Argument<List<FileSystemInfo>>("input-files"));
            command.Add(new Option<string>(new string[] {"-p", "--input-pattern"}, "Input file name pattern"));
            command.Add(new Option<FileInfo>(new string[] {"-l", "--version-list"}, "Version list file"));
            command.Add(new Option<string>(new string[] {"-r", "--current-version"}, "Current version"));
            command.Add(new Option<string>(new string[] {"-o", "--output-dir"}, "Output directory"));
            command.Add(new Option<string>(new string[] {"-n", "--newline"}, "Newline code(cr,lf,crlf)"));
            command.Add(new Option<FileInfo>(new string[] {"-c", "--config-file"}, "Config file"));

            command.Handler = CommandHandler.Create(
                (List<FileSystemInfo> inputFiles,
                string inputPattern,
                FileInfo versionList,
                string currentVersion,
                string outputDir,
                string newLine,
                FileInfo configFile) =>
                {
                    var config = new Config();
                    config.ConfigFile = configFile;
                    config.LoadFromFile();

                    if (inputFiles != null && inputFiles.Count > 0) config.InputFiles = inputFiles;
                    if (!String.IsNullOrEmpty(inputPattern)) config.InputPattern = inputPattern;
                    if (versionList != null) config.VersionList = versionList;
                    if (!String.IsNullOrEmpty(currentVersion)) config.CurrentVersion = currentVersion;
                    if (!String.IsNullOrEmpty(outputDir)) config.OutputDir = outputDir;
                    if (!String.IsNullOrEmpty(newLine)) config.NewLine = newLine;
                        
                    Convert(config);
                });
            
            command.Invoke(args);
        }

        static void Convert(Config config)
        {
            if (!Directory.Exists(config.OutputDir))
            {
                Console.WriteLine($"Output directory [{config.OutputDir}] does not exist.");
                return;
            }
            
            var regex = new Regex(config.InputPattern);
            var excludeRegex = new Regex(config.ExcludeInputPattern);
            
            var tasks = config.InputFiles.SelectMany(info =>
            {
                if ((!info.Exists))
                {
                    Console.WriteLine($"Input file/directory [{info.FullName}] does not exist.");
                    return new string[]{};
                }
                else if ((info.Attributes & FileAttributes.Directory) == FileAttributes.Directory)
                {
                    return Directory.GetFiles(info.FullName, "*", SearchOption.AllDirectories);
                }
                else
                {
                    return new string[] {info.FullName};
                }
            }).Where(path =>
            {
                var fileName = Path.GetFileName(path);
                return regex.IsMatch(fileName) && !excludeRegex.IsMatch(fileName);
            }).Distinct().SelectMany(path =>
            {
                var book = WorkbookFactory.Create(path, password: null, readOnly: true);
                return book.GetSheetEnumerable();
            }).Select(sheet =>
            {
                return Task.Factory.StartNew(() =>
                {
                    var converter = new Converter(config);
                    (var name, var outputs) = converter.Convert(sheet);
                    if (String.IsNullOrEmpty(name))
                    {
                        Console.WriteLine($"{sheet.SheetName}: invalid format");
                        return;
                    }
                    Console.WriteLine($"Convert: {name}");
                    TextWriter.Write(name, outputs, config);
                });
            });

            Task.WaitAll(tasks.ToArray());
        }
        
    }
}
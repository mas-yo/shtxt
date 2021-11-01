using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using NPOI.SS.UserModel;

namespace shtxt
{
    public class Converter
    {
        public static (string, IEnumerable<IReadOnlyCollection<string>>) Convert(ISheet sheet, Config config)
        {
            var controlParser = new ControlParser() { CommentStartsWith = config.CommentStartsWith };
            var loader = new SheetLoader(config.TableNameTag, config.ColumnControlTag, config.ColumnNameTag, controlParser);
            var info = loader.Load(sheet.GetRowDataEnumerable());
            if (!info.IsValid) return (null, null);
                
            var versionList = new List<string>();
            if (File.Exists(config.VersionList.FullName))
            {
                versionList = File.ReadLines(config.VersionList.FullName).ToList();
            }

            return (info.Header.Name, GetOutputRowEnumerable(info, versionList, config.CurrentVersion, config.OutputColumnNameTag));
        }
        
        static IEnumerable<IReadOnlyCollection<string>> GetOutputRowEnumerable(SheetInfo info, IList<string> versionList, string currentVersion, string outputColumnNameTag)
        {
            var columnInfos = info.GetEnumerableColumnInfo();
            var skipColumns = columnInfos
                .Select((columnInfo, idx) => (idx, IsEnable(columnInfo.Control, currentVersion, versionList)))
                .Where(i => i.Item2 == false)
                .Select(i => i.idx)
                .ToList();

            var columnNames = columnInfos
                .Where((ci, idx) => !skipColumns.Contains(idx))
                .Select(ci => ci.Name)
                .ToList();

            if (!String.IsNullOrEmpty(outputColumnNameTag))
            {
                columnNames.Insert(0, outputColumnNameTag);
            }

            yield return columnNames.AsReadOnly();

                foreach (var row in info.Body)
            {
                if (!IsEnable(row.Control, currentVersion, versionList)) continue;
                
                yield return row.Data
                    .Where((data, idx) => !skipColumns.Contains(idx))
                    .ToList()
                    .AsReadOnly();
            }
        }
        
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


    }
}
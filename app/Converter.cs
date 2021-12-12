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
        private Config config;
        private IList<string> versionList;

        public Converter(Config config)
        {
            this.config = config;
            versionList = new List<string>();
            if (File.Exists(config.VersionList.FullName))
            {
                versionList = File.ReadLines(config.VersionList.FullName).ToList();
            }
        }
        
        public IEnumerable<IReadOnlyCollection<string>> GetOutputRowEnumerable(SheetInfo info)
        {
            var columnInfos = info.GetEnumerableColumnInfo();
            var skipColumns = columnInfos
                .Select((columnInfo, idx) => (idx, IsEnable(columnInfo.Control)))
                .Where(i => i.Item2 == false)
                .Select(i => i.idx)
                .ToList();

            var columnNames = columnInfos
                .Where((ci, idx) => !skipColumns.Contains(idx))
                .Select(ci => ci.Name)
                .ToList();

            if (config.IsOutputControlColumn)
            {
                columnNames.Insert(0, config.OutputColumnNameTag);
            }

            yield return columnNames.AsReadOnly();

            foreach (var row in info.Body)
            {
                if (!IsEnable(row.Control)) continue;

                var list = row.Data
                    .Where((data, idx) => !skipColumns.Contains(idx))
                    .Select(data => data.Replace("\n", "\\n"))
                    .ToList();
                
                if (config.IsOutputControlColumn)
                {
                    list.Insert(0, GetControlTag(row.Control));
                }

                yield return list.AsReadOnly();
            }
        }
        
        private bool IsEnable(Control control)
        {
            switch (control)
            {
                case None:
                    return true;
                case Comment:
                    if (config.IsOutputControlColumn)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                    
                case Version(var comp, var version):
                    var currentIndex = versionList.IndexOf(config.CurrentVersion);
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

        private string GetControlTag(Control control)
        {
            switch (control)
            {
                case Comment:
                    return config.OutputCommentTag;
            }
            return "";
        }


    }
}
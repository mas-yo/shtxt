using System;
using System.Collections.Generic;
using System.Linq;


namespace shtxt
{
    using ColumnIndex = Int32;
    using RowEnumerator = IEnumerator<IDictionary<int, string>>;
    
    public class SheetLoader
    {
        const int CONTROL_COLUMN = 0;

        
        public static List<string> DictionaryToList(IDictionary<ColumnIndex, string> dict)
        {
            var list = new List<string>();
            int max = dict.Keys.Max();
            for (var i = CONTROL_COLUMN + 1; i <= max; i++)
            {
                if (dict.ContainsKey(i))
                {
                    list.Add(dict[i]);
                }
                else
                {
                    list.Add("");
                }
            }

            return list;
        }

        private string tableNameTag;
        private string columnCommandTag;
        private string columnNameTag;
        private ControlParser controlParser;
        private RowEnumerator rowEnumerator;

        public SheetLoader(string tableNameTag, string columnCommandTag, string columnNameTag, ControlParser controlParser)
        {
            this.tableNameTag = tableNameTag;
            this.columnCommandTag = columnCommandTag;
            this.columnNameTag = columnNameTag;
            this.controlParser = controlParser;
        }

        private SheetInfo.HeaderInfo LoadHeaderInfo()
        {
            string tableName = null;
            List<string> columnCommands = null;
            List<string> columnNames = null;

            while(rowEnumerator.MoveNext())
            {
                var columns = rowEnumerator.Current;
                if (columns.Count <= 0)
                    continue;

                if (columns.ContainsKey(CONTROL_COLUMN))
                {
                    var control = columns[CONTROL_COLUMN];

                    if (control == tableNameTag && columns.ContainsKey(CONTROL_COLUMN + 1))
                    {
                        tableName = columns[CONTROL_COLUMN + 1];
                    }
                    else if (control == columnCommandTag)
                    {
                        columnCommands = DictionaryToList(columns);
                    }
                    else if (control == columnNameTag)
                    {
                        columnNames = DictionaryToList(columns);
                        break;
                    }
                }
            }

            var controls = columnCommands != null ? columnCommands.Select(c => controlParser.Parse(c)).ToList() : null;

            var errorMessages = new List<string>();
            if (String.IsNullOrEmpty(tableName))
            {
                errorMessages.Add($"can not detect table name. {tableNameTag} may not exists, or table name not specified");
            }

            if (columnNames == null)
            {
                errorMessages.Add($"{columnNameTag} not exists");
            }
            else if (columnNames.Count <= 0)
            {
                errorMessages.Add($"column name not specified");
            }

            return new SheetInfo.HeaderInfo(tableName, controls, columnNames, errorMessages);
        }

        private IReadOnlyList<SheetInfo.Row> LoadBody(int columnCount)
        {
            var body = new List<SheetInfo.Row>();
            while(rowEnumerator.MoveNext())
            {
                var columns = rowEnumerator.Current;
                if (columns.Values.All(x => x == "")) continue;
                if (columns.Count == 0) continue;
                
                string control = null;
                
                if (columns.ContainsKey(CONTROL_COLUMN))
                {
                    control = columns[CONTROL_COLUMN];
                }

                var list = DictionaryToList(columns);
                if (list.Count < columnCount)
                {
                    list.AddRange(Enumerable.Repeat("", columnCount - list.Count));
                }
                else if (list.Count > columnCount)
                {
                    list.RemoveRange(columnCount, list.Count - columnCount);
                }
                body.Add(new SheetInfo.Row(controlParser.Parse(control), list));
            }

            return body;
        }
        
        public SheetInfo Load(IEnumerable<IDictionary<ColumnIndex, string>> rows)
        {
            rowEnumerator = rows.GetEnumerator();
            var header = LoadHeaderInfo();
            if (!header.IsValid)
            {
                return new SheetInfo(header, null);
            }

            var body = LoadBody(header.ColumnNames.Count);
            return new SheetInfo(header, body);
        }
    }
}
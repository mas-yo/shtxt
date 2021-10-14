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
        private RowEnumerator rowEnumerator;

        public SheetLoader(string tableNameTag, string columnCommandTag, string columnNameTag)
        {
            this.tableNameTag = tableNameTag;
            this.columnCommandTag = columnCommandTag;
            this.columnNameTag = columnNameTag;
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

            return new SheetInfo.HeaderInfo(tableName, columnCommands, columnNames);
        }

        private IReadOnlyList<SheetInfo.Row> LoadBody()
        {
            var body = new List<SheetInfo.Row>();
            while(rowEnumerator.MoveNext())
            {
                var columns = rowEnumerator.Current;
                string control = null;
                
                if (columns.ContainsKey(CONTROL_COLUMN))
                {
                    control = columns[CONTROL_COLUMN];
                }

                body.Add(new SheetInfo.Row(control, DictionaryToList(columns)));
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

            var body = LoadBody();
            return new SheetInfo(header, body);
        }
    }
}
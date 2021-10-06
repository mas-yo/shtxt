using System;
using System.Collections.Generic;
using System.Linq;


namespace MasterDataConverter
{
    using ColumnIndex = Int32;
    using RowEnumerator = IEnumerator<IDictionary<int, string>>;
    
    public class MasterTableLoader
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

        public MasterTableLoader(string tableNameTag, string columnCommandTag, string columnNameTag)
        {
            this.tableNameTag = tableNameTag;
            this.columnCommandTag = columnCommandTag;
            this.columnNameTag = columnNameTag;
        }

        private MasterTableInfo.HeaderInfo LoadHeaderInfo()
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

            return new MasterTableInfo.HeaderInfo(tableName, columnCommands, columnNames);
        }

        private IReadOnlyList<MasterTableInfo.Row> LoadBody()
        {
            var body = new List<MasterTableInfo.Row>();
            while(rowEnumerator.MoveNext())
            {
                var columns = rowEnumerator.Current;
                string control = null;
                
                if (columns.ContainsKey(CONTROL_COLUMN))
                {
                    control = columns[CONTROL_COLUMN];
                }

                body.Add(new MasterTableInfo.Row(control, DictionaryToList(columns)));
            }

            return body;
        }
        
        public MasterTableInfo Load(IEnumerable<IDictionary<ColumnIndex, string>> rows)
        {
            rowEnumerator = rows.GetEnumerator();
            var header = LoadHeaderInfo();
            if (!header.IsValid)
            {
                return new MasterTableInfo(header, null);
            }

            var body = LoadBody();
            return new MasterTableInfo(header, body);
        }
    }
}
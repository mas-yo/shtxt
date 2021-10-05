using System;
using System.Collections.Generic;
using System.Linq;


namespace MasterDataConverter
{
    using ColumnIndex = Int32;


    class BodyRowInfo
    {
        public string Command {get;set;}
        public IList<string> Data {get;set;}
    }

    class MasterTableInfo
    {
        const int COMMAND_COLUMN = 0;

        public string Name{get; private set;}
        public IList<string> ColumnCommands;
        public IList<string> ColumnNames{get; private set;}
        public IList<BodyRowInfo> Body { get; private set; } = new List<BodyRowInfo>();

        public static IList<string> GetDataList(IDictionary<int, string> dict)
        {
            var list = new List<string>();
            int max = dict.Keys.Max();
            for (var i = COMMAND_COLUMN + 1; i <= max; i++)
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

        public static void ExtendList(IList<string> list, int count)
        {
            int start = list.Count;
            for (int i = start; i < count; i++)
            {
                list.Add("");
            }
        }
        public static MasterTableInfo Load(IEnumerable<IDictionary<ColumnIndex, string>> rows)
        {
            
            var info = new MasterTableInfo();

            foreach (var columns in rows)
            {
                if (columns.Count <= 0) continue;

                if (columns.ContainsKey(COMMAND_COLUMN))
                {
                    var command = columns[COMMAND_COLUMN];
                    if (command.StartsWith("#")) continue;

                    if (command == "[テーブル名]" && columns.ContainsKey(COMMAND_COLUMN + 1))
                    {
                        info.Name = columns[COMMAND_COLUMN + 1];
                    }
                    else if (command == "[カラム名]")
                    {
                        info.ColumnNames = GetDataList(columns);
                    }
                    else if (command == "[カラム制御]")
                    {
                        info.ColumnCommands = GetDataList(columns);
                    }
                }
                
                if (info.ColumnNames != null)
                {
                    var row = new BodyRowInfo();
                    if (columns.ContainsKey(COMMAND_COLUMN)) 
                    {
                        row.Command = columns[COMMAND_COLUMN];
                    }

                    row.Data = GetDataList(columns);
                    info.Body.Add(row);
                }
            }

            int max = info.CalcMaxColumnCount();
            ExtendList(info.ColumnNames, max);
            foreach (var row in info.Body)
            {
                ExtendList(row.Data, max);
            }

            return info;
        }

        private int CalcMaxColumnCount()
        {
            int columnMax = ColumnNames.Count;
            int bodyMax = Body.Max(r => r.Data.Count);
            return Math.Max(columnMax, bodyMax);
        }

        public bool IsValid
        {
            get {
                if (String.IsNullOrWhiteSpace(Name)) return false;
                if (ColumnNames == null) return false;
                if (ColumnNames.Count == 0) return false;
                if (Body == null) return false;
                if (ColumnNames.Any(c => String.IsNullOrWhiteSpace(c))) return false;
                return true;
            }
        }
    }

}

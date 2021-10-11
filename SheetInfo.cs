using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;

using NPOI.SS.Formula.Functions;


namespace Shtxt
{
    using ColumnIndex = Int32;



    public class SheetInfo
    {
        const int CONTROL_COLUMN = 0;

        public class HeaderInfo
        {
            public string Name { get; private set; }
            public IReadOnlyCollection<string> ColumnCommands { get; private set; }
            public IReadOnlyCollection<string> ColumnNames { get; private set; }

            public HeaderInfo(string name, IReadOnlyCollection<string> columnCommands,
                IReadOnlyCollection<string> columnNames)
            {
                Name = name;
                ColumnCommands = columnCommands;
                ColumnNames = columnNames;
            }

            public bool IsValid
            {
                get
                {
                    if (String.IsNullOrWhiteSpace(Name)) return false;
                    if (ColumnNames == null) return false;
                    if (ColumnNames.Count == 0) return false;
                    if (ColumnNames.Any(c => String.IsNullOrWhiteSpace(c)))
                        return false;
                    return true;
                }
            }
        }
        public class Row
        {
            public string Control { get; set; }
            public IReadOnlyCollection<string> Data { get; private set; }

            public Row(string control, IReadOnlyCollection<string> data)
            {
                Control = control;
                Data = data;
            }
        }


        public HeaderInfo Header { get; private set; }
        public IReadOnlyList<Row> Body { get; private set; }

        public SheetInfo(HeaderInfo header, IReadOnlyList<Row> body)
        {
            Header = header;
            Body = body;
        }

        private int CalcMaxColumnCount()
        {
            int columnMax = Header.ColumnNames.Count;
            int bodyMax = Body.Max(r => r.Data.Count);
            return Math.Max(columnMax, bodyMax);
        }

        public bool IsValid
        {
            get
            {
                if (!Header.IsValid) 
                    return false;
                
                if (Body == null)
                    return false;
                if (Body.Count == 0)
                    return false;
                return true;
            }
        }
    }

}
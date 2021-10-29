using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Security.Cryptography.X509Certificates;

using NPOI.HSSF.Record;
using NPOI.SS.Formula.Functions;

using Org.BouncyCastle.Asn1;


namespace shtxt
{
    using ColumnIndex = Int32;
    
    public class SheetInfo
    {
        const int CONTROL_COLUMN = 0;

        public class HeaderInfo
        {
            public string Name { get; private set; }
            public IReadOnlyCollection<Control> ColumnControls { get; private set; }
            public IReadOnlyCollection<string> ColumnNames { get; private set; }

            public HeaderInfo(string name, IReadOnlyCollection<Control> columnControls,
                IReadOnlyCollection<string> columnNames)
            {
                Name = name;
                ColumnControls = columnControls;
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


        public class ColumnInfo
        {
            public Control Control { get; set; }
            public string Name { get; set; } = "";
        }
        public class Row
        {
            public Control Control { get; private set; }
            public IReadOnlyCollection<string> Data { get; private set; }

            public Row(Control control, IReadOnlyCollection<string> data)
            {
                Control = control;
                Data = data;
            }
        }


        public HeaderInfo Header { get; private set; }
        public IReadOnlyCollection<Row> Body { get; private set; }

        public SheetInfo(HeaderInfo header, IReadOnlyList<Row> body)
        {
            Header = header;
            Body = body;
        }

        public IEnumerable<ColumnInfo> GetEnumerableColumnInfo()
        {
            return Header.ColumnControls.Zip(Header.ColumnNames, 
                (control,name) => new ColumnInfo(){Control = control, Name = name} );
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
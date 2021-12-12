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

            private List<Control> _columnControl;
            public IReadOnlyCollection<Control> ColumnControls { get => _columnControl; }
            public IReadOnlyCollection<string> ColumnNames { get; private set; }
            
            public IReadOnlyCollection<string> ErrorMessages { get; private set; }

            public HeaderInfo(string name, List<Control> columnControls,
                IReadOnlyCollection<string> columnNames, IReadOnlyCollection<string> errorMessages)
            {
                ErrorMessages = errorMessages;
                if (String.IsNullOrEmpty(name)) return;
                if (columnNames == null) return;
                
                Name = name;
                ColumnNames = columnNames;

                _columnControl = columnControls;
                if (_columnControl == null)
                {
                    _columnControl = ColumnNames.Select(_ => (Control)new None()).ToList();
                }
                else if (ColumnControls.Count < columnNames.Count)
                {
                    _columnControl.Add(new None());
                }
            }

            public bool IsValid => ErrorMessages == null || ErrorMessages.Count == 0;
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
    }

}
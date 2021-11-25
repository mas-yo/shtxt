using System.Collections.Generic;

using NPOI.SS.UserModel;

namespace shtxt
{
    public static class PoiExtension
    {
        public static IEnumerable<ISheet> GetSheetEnumerable(this IWorkbook book)
        {
            var enm = book.GetEnumerator();
            while (enm.MoveNext())
            {
                yield return enm.Current;
            }
        }
        public static IEnumerable<IDictionary<int, string>> GetRowDataEnumerable(this ISheet sheet)
        {
            var enm = sheet.GetRowEnumerator();
            while (enm.MoveNext())
            {
                var row = (IRow)enm.Current;
                var data = new Dictionary<int, string>();
                foreach (var cell in row.Cells)
                {
                    switch (cell.CellType)
                    {
                        case CellType.Numeric:
                            data.Add(cell.ColumnIndex, cell.NumericCellValue.ToString());
                            break;
                        case CellType.String:
                            data.Add(cell.ColumnIndex, cell.StringCellValue);
                            break;
                        case CellType.Boolean:
                            data.Add(cell.ColumnIndex, cell.BooleanCellValue.ToString());
                            break;
                        case CellType.Formula:
                            switch (cell.CachedFormulaResultType)
                            {
                                case CellType.Numeric:
                                    data.Add(cell.ColumnIndex, cell.NumericCellValue.ToString());
                                    break;
                                case CellType.String:
                                    data.Add(cell.ColumnIndex, cell.StringCellValue);
                                    break;
                            }
                            break;
                    }
                }

                yield return data;
            }
        }        
    }
}
using System;
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

        public static string GetText(this ICell cell, string dateTimeFormat)
        {
            var cellType = cell.CellType;
            if (cellType == CellType.Formula)
            {
                cellType = cell.CachedFormulaResultType;
            }
            switch (cellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue.ToString(dateTimeFormat);
                    }
                    else
                    {
                        return cell.NumericCellValue.ToString();
                    }
                
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                
                case CellType.Blank:
                    return cell.ToString();
                
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                
                case CellType.Unknown:
                    Console.WriteLine("unknown cell type detected");
                    return "";
            }

            return "";
        }
        public static IEnumerable<IDictionary<int, string>> GetRowDataEnumerable(this ISheet sheet, string dateTimeFormat)
        {
            var enm = sheet.GetRowEnumerator();
            while (enm.MoveNext())
            {
                var row = (IRow)enm.Current;
                var data = new Dictionary<int, string>();
                foreach (var cell in row.Cells)
                {
                    data.Add(cell.ColumnIndex, cell.GetText(dateTimeFormat));
                }

                yield return data;
            }
        }        
    }
}
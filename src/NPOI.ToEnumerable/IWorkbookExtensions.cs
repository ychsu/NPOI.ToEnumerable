using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.ToEnumerable.Exceptions;

namespace NPOI.ToEnumerable
{
    public static class IWorkbookExtensions
    {
        public static IEnumerable<T> ToEnumerable<T>(this IWorkbook workbook,
            int idx)
            where T: class, new()
        {
            var sheet = workbook.GetSheetAt(idx);
            if (sheet == null)
            {
                throw new WorkSheetNotExistsException(idx);
            }
            return sheet.ToEnumerable<T>();
        }

        public static IEnumerable<T> ToEnumerable<T>(this IWorkbook workbook,
            string sheetName)
            where T: class, new()
        {
            var sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                throw new WorkSheetNotExistsException(sheetName);
            }
            return sheet.ToEnumerable<T>();
        }

        public static IEnumerable<T> ToEnumerable<T>(this IWorkbook workbook,
            params int[] sheetIndexes)
            where T: class, new()
        {
            return sheetIndexes
                .SelectMany(idx => ToEnumerable<T>(workbook, idx));
        }

		public static IEnumerable<T> ToEnumerable<T>(this IWorkbook workbook)
			where T: class, new()
		{
			var sheetCount = workbook.NumberOfSheets;
			return ToEnumerable<T>(workbook, Enumerable.Range(0, sheetCount).ToArray());
		}
    }
}

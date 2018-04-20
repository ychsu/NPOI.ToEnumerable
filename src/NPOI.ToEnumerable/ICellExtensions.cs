using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

namespace NPOI.ToEnumerable
{
	internal static class ICellExtensions
	{
		public static string GetString(this ICell cell)
		{
			if (cell == null)
			{
				return null;
			}
			if (cell.CellType == CellType.String)
			{
				return cell.StringCellValue;
			}
			return new DataFormatter(CultureInfo.CurrentCulture).FormatCellValue(cell);
		}

		public static object GetValue(this ICell cell)
		{
			if (cell == null)
			{
				return null;
			}
			object value = null;
			switch (cell.CellType)
			{
				case CellType.Numeric:
					value = DateUtil.IsCellDateFormatted(cell) == true ?
						(object)cell.DateCellValue :
						(object)cell.NumericCellValue;
					break;
				case CellType.String:
					value = cell.StringCellValue;
					break;
				case CellType.Formula:
					throw new NotImplementedException();
				case CellType.Boolean:
					value = (object)cell.BooleanCellValue;
					break;
				case CellType.Blank:
				case CellType.Error:
				case CellType.Unknown:
				default:
					break;
			}
			return value;
		}

		public static T GetValue<T>(this ICell cell)
		{
			var type = typeof(T);
			var value = GetValue(cell);
			var parse = type.GetMethod("Parse", new Type[] { typeof(string) });
			var str = value?.ToString();
			if (type == typeof(string))
			{
				return (T)(object)str;
			}
			if (value == null || value.GetType() == type || parse == null)
			{
				return (T)value;
			}
			return (T)parse.Invoke(null, new object[] { str });
		}

		public static T GetEnum<T>(this ICell cell)
			where T : struct
		{
			var type = typeof(T);
			var value = cell?.GetString().Trim();
			T result;
			if (Enum.TryParse(value, true, out result) == false)
			{
				string ex = string.Format("{0} RowIndex:{1} ColumnIndex:{2}", new ArgumentOutOfRangeException().Message, cell.RowIndex, cell.ColumnIndex);
				throw new Exception();
			}
			return result;
		}
	}
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

namespace NPOI.ToEnumerable
{
	internal static class IRowExtensions
	{
		public static T ToItem<T>(this IRow row, IDictionary<int, Action<T, ICell>> actions)
			where T : class, new()
		{
			T item = new T();
			foreach (var action in actions)
			{
				var cell = row.Cells[action.Key];
				action.Value.Invoke(item, cell);
			}
			return item;
		}
	}
}

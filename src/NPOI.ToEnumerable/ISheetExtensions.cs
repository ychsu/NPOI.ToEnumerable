using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;

namespace NPOI.ToEnumerable
{
	public static class ISheetExtensions
	{
		public static IEnumerable<T> ToEnumerable<T>(this ISheet sheet)
			where T : class, new()
		{
			var enumerator = sheet.GetRowEnumerator();
			if (enumerator.MoveNext() == false)
			{
				yield break;
			}
			var headerRow = enumerator.Current as IRow;
			// todo: 透過header row 決定要怎麼建立一個T
			var actions = GetActions<T>(headerRow);
			while (enumerator.MoveNext() == true)
			{
				var row = enumerator.Current as IRow;
				yield return row.ToItem<T>(actions);
			}
		}

		private static IDictionary<int, Action<T, ICell>> GetActions<T>(IRow headerRow)
		{
			var type = typeof(T);

			var columns = headerRow.Cells.Select((item, idx) => new
			{
				idx,
				item = item.StringCellValue
			});
			return type.GetProperties()
				.Select(p => new
				{
					DisplayName = p.GetCustomAttribute<DisplayAttribute>(true)?.GetName() ?? p.Name,
					Prop = p
				})
				.Join(columns, p => p.DisplayName, q => q.item, (p, q) => new
				{
					q.idx,
					p.Prop
				})
				.ToDictionary(p => p.idx, q => GetAssignAction<T>(q.Prop));
		}

		private static Action<T, ICell> GetAssignAction<T>(PropertyInfo prop)
		{
			var arg1 = Expression.Parameter(typeof(T), "arg1");
			var arg2 = Expression.Parameter(typeof(ICell), "arg2");
			var member = Expression.PropertyOrField(arg1, prop.Name);
			var method = GetGetValueMethod().MakeGenericMethod(prop.PropertyType);
			var callExp = Expression.Call(method, arg2);

			var setter = Expression.Lambda<Action<T, ICell>>(
				Expression.Assign(member, callExp),
				arg1, 
				arg2);

			return setter.Compile();
		}

		private static MethodInfo getValueMethod;
		private static MethodInfo GetGetValueMethod()
		{
			if (getValueMethod == null)
			{
				getValueMethod = typeof(ICellExtensions)
					.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic)
					.FirstOrDefault(p => p.Name == "GetValue" && p.IsGenericMethod);
			}
			return getValueMethod;
		}
	}
}

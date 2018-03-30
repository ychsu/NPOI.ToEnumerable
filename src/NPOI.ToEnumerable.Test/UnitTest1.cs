using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.DumpExcel;
using NPOI.SS.UserModel;
using NPOI.ToEnumerable.Test.Models;

namespace NPOI.ToEnumerable.Test
{
	[TestClass]
	public class UnitTest1
	{
		[TestMethod]
		public void TestMethod1()
		{
			var enumerable = Enumerable.Range(1, 10)
				.Select(p => new Foo
				{
					DT = DateTimeOffset.Now.AddDays(p),
					DT2 = DateTime.Now.AddDays(p),
					Name = $"Foo{p}",
					SerId = p
				});

			var workbook = enumerable.DumpXLS();

			var result = workbook.ToEnumerable<Foo>(0);

			Assert.AreEqual(enumerable.Count(), result.Count());
		}
	}
}

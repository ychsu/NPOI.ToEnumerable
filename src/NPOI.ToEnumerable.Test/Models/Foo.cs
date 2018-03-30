using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.DumpExcel.Attributes;

namespace NPOI.ToEnumerable.Test.Models
{

	[Sheet(Name = "Foo Sheet")]
	public class Foo
	{
		[ExcelColumn(Format: "#.00", Width: 20)]
		public int SerId { get; set; }

		[ExcelColumn]
		public string Name { get; set; }

		[ExcelColumn(Format: "yyyy-MM-dd", Width: 20)]
		public DateTimeOffset DT { get; set; }

		[ExcelColumn(Format: "yyyy-MM-dd", Width: 20)]
		public DateTime DT2 { get; set; } = DateTime.Now;
	}
}

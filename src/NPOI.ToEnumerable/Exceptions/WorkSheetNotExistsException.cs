using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.ToEnumerable.Exceptions
{
    public class WorkSheetNotExistsException
        : Exception
    {
        public WorkSheetNotExistsException(string sheetName)
            :base($"{sheetName} 不存在workbook中")
        {

        }

        public WorkSheetNotExistsException(int idx)
            : base ($"workbook 中沒有第 {idx} 個sheet")
        {

        }
    }
}

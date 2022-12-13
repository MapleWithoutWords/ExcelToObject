using OEM.Core.Dto;
using System;
using System.Collections.Generic;
using System.Text;

namespace OEM.Core.ExcelOperator
{
    public interface IExcelBaseService : IDisposable
    {
        public (List<ExcelNameManagerOutput>, int) GetNameManagerList(string sheetName);
        public object GetWorkbook();
    }
}

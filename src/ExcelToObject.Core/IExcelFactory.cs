using OEM.Core.ExcelOperator;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OEM.Core
{
    public interface IExcelFactory
    {
        public IExcelAppService Create(Stream excelStream);
        public IExcelAppService Create(string filePath);
    }
}

using OEM.Core;
using OEM.Core.ExcelOperator;
using OEM.Npoi.ExcelOperatorImpl;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OEM.Npoi
{
    public class NpoiExcelFactory : IExcelFactory
    {
        public IExcelAppService Create(Stream excelStream)
        {
            return new NpoiExcelAppService(excelStream);
        }

        public IExcelAppService Create(string filePath)
        {
            return new NpoiExcelAppService(filePath);
        }
    }
}

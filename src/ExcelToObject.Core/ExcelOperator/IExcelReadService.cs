using System;
using System.Collections.Generic;
using System.Text;

namespace OEM.Core.ExcelOperator
{
    public interface IExcelReadService : IExcelBaseService
    {

        /// <summary>
        /// 根据名称管理器读取数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public T ReadByNameManager<T>(string sheetName = "Sheet1") where T : class, new();
        /// <summary>
        /// 根据名称管理器读取数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public List<T> ReadListByNameManager<T>(string sheetName = "Sheet1") where T : class, new();
    }
}

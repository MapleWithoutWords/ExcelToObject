using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OEM.Core.ExcelOperator
{
    public interface IExcelWriteService : IExcelBaseService
    {
        /// <summary>
        /// 根据名称管理器写入数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="datas"></param>
        /// <param name="sheetName"></param>
        public void WriteListByNameManager<T>(List<T> datas, string sheetName = "Sheet1") where T : class, new();

        /// <summary>
        /// 根据名称管理器写入数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="sheetName"></param>
        public void WriteByNameManager<T>(T data, string sheetName = "Sheet1") where T : class, new();

        /// <summary>
        /// 写入到新文件
        /// </summary>
        /// <param name="filePath"></param>
        public void Write(string filePath);
        /// <summary>
        /// 写入到流中
        /// </summary>
        /// <param name="fileStream"></param>
        public void Write(Stream fileStream);
    }
}

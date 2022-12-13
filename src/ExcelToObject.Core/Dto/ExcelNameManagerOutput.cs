using System;

namespace OEM.Core.Dto
{
    /// <summary>
    /// excel名称管理器dto
    /// </summary>
    public class ExcelNameManagerOutput
    {
        public ExcelNameManagerOutput(string name, int rowIndex, int columnIndex)
        {
            Name = name;
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }
        /// <summary>
        /// 名称管理器的值
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 列下标
        /// </summary>
        public int ColumnIndex { get; set; }
        /// <summary>
        /// 行下标
        /// </summary>
        public int RowIndex { get; set; }
    }
}

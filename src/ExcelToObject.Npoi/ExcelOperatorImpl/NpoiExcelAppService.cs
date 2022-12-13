using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using OEM.Core.Dto;
using OEM.Core.ExcelOperator;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OEM.Npoi.ExcelOperatorImpl
{
    public class NpoiExcelAppService : IExcelAppService
    {
        public IDictionary<string, (List<ExcelNameManagerOutput>, int)> CacheNameManagerDic { get; set; }
        protected IWorkbook Workbook { get; set; }

        public NpoiExcelAppService(Stream excelStream)
        {
            Workbook = WorkbookFactory.Create(excelStream);
        }
        public NpoiExcelAppService(string filePath)
        {
            Workbook = WorkbookFactory.Create(filePath);
        }

        public void Dispose()
        {
            Workbook?.Close();
            Workbook?.Dispose();
            GC.Collect();
        }

        public (List<ExcelNameManagerOutput>, int) GetNameManagerList(string sheetName)
        {
            if (CacheNameManagerDic == null)
            {
                InitNameManager();
            }
            return CacheNameManagerDic[sheetName];
        }

        private void InitNameManager()
        {
            CacheNameManagerDic = new Dictionary<string, (List<ExcelNameManagerOutput>, int)>();
            for (int i = 0; i < Workbook.NumberOfNames; i++)
            {
                var nameObj = Workbook.GetNameAt(i);

                int.TryParse(nameObj.RefersToFormula.Substring(nameObj.RefersToFormula.LastIndexOf("$") + 1), out int rowIndex);

                //计算列坐标
                var columnRightStr = nameObj.RefersToFormula.Substring(nameObj.RefersToFormula.IndexOf("$") + 1);
                var columnLetter = columnRightStr.Substring(0, columnRightStr.IndexOf("$"));
                var columnIndex = GetExcelColumnLetterNumber(columnLetter);

                if (CacheNameManagerDic.ContainsKey(nameObj.SheetName) == false)
                {
                    CacheNameManagerDic[nameObj.SheetName] = (new List<ExcelNameManagerOutput>(), rowIndex);
                }
                var nameManagerItem = CacheNameManagerDic[nameObj.SheetName];
                nameManagerItem.Item1.Add(new ExcelNameManagerOutput(nameObj.NameName, rowIndex, columnIndex));
                nameManagerItem.Item2 = rowIndex;

            }
        }

        /// <summary>
        /// 根据名称管理器的值，计算所在列坐标
        /// </summary>
        /// <param name="columnLetter"></param>
        /// <returns></returns>
        private static int GetExcelColumnLetterNumber(string columnLetter)
        {
            var columnIndex = 0;
            var powNum = columnLetter.Length - 1;
            foreach (var item in columnLetter)
            {
                var itemNum = item - 65;
                var baseNum = (int)Math.Pow(26, powNum);
                if (powNum < 1)
                {
                    columnIndex += itemNum;
                }
                else
                {
                    columnIndex += (itemNum + 1) * baseNum;
                }
                powNum--;
            }

            return columnIndex;
        }


        public T ReadByNameManager<T>(string sheetName = "Sheet1") where T : class, new()
        {
            var sheet = Workbook.GetSheet(sheetName);

            var (nameManagerList, startRowIndex) = GetNameManagerList(sheetName);

            var objProperties = typeof(T).GetProperties();
            T obj = new T();
            foreach (var item in objProperties)
            {
                var nameManger = nameManagerList.FirstOrDefault(e => e.Name == item.Name);
                if (nameManger == null)
                {
                    continue;
                }
                var row = sheet.GetRow(nameManger.RowIndex - 1);
                var cell = row.GetCell(nameManger.ColumnIndex);
                if (cell == null)
                {
                    object value = GetDefaultValue(item.PropertyType);
                    item.SetValue(obj, value);
                    continue;
                }
                try
                {
                    object value = Convert(cell, item.PropertyType);
                    item.SetValue(obj, value);
                }
                catch (Exception ex)
                {
                    throw new ApplicationException($"excel中，[{nameManger.Name}]第{nameManger.RowIndex + 1}行，第{nameManger.ColumnIndex + 1}列，数据格式非法,请检查】");
                }

            }
            return obj;
        }

        public List<T> ReadListByNameManager<T>(string sheetName = "Sheet1") where T : class, new()
        {
            var sheet = Workbook.GetSheet(sheetName);

            var (nameManagerList, startRowIndex) = GetNameManagerList(sheetName);

            var objProperties = typeof(T).GetProperties();
            List<T> result = new List<T>();

            for (int i = startRowIndex; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null)
                {
                    continue;
                }
                if (row.GetCell(0) == null || row.GetCell(0).CellType == CellType.Blank)
                {
                    continue;
                }

                T obj = new T();
                foreach (var item in objProperties)
                {
                    var nameManger = nameManagerList.FirstOrDefault(e => e.Name == item.Name);
                    if (nameManger == null)
                    {
                        continue;
                    }

                    var cell = row.GetCell(nameManger.ColumnIndex);
                    if (cell == null)
                    {
                        object value = GetDefaultValue(item.PropertyType);
                        item.SetValue(obj, value);
                        continue;
                    }
                    try
                    {
                        object value = Convert(cell, item.PropertyType);
                        item.SetValue(obj, value);
                    }
                    catch (Exception ex)
                    {
                        throw new ApplicationException($"excel中，[{nameManger.Name}]第{i + 1}行，第{nameManger.ColumnIndex + 1}列，数据格式非法,请检查】");
                    }

                }

                result.Add(obj);

            }

            return result;
        }

        public void WriteByNameManager<T>(T data, string sheetName = "Sheet1") where T : class, new()
        {
            var sheet = Workbook.GetSheet(sheetName);

            var (nameManagerList, startRowIndex) = GetNameManagerList(sheetName);

            var objProperties = typeof(T).GetProperties();
            foreach (var item in objProperties)
            {
                var value = item.GetValue(data);
                if (value == null)
                {
                    continue;
                }
                var nameManger = nameManagerList.FirstOrDefault(e => e.Name == item.Name);
                if (nameManger == null)
                {
                    continue;
                }
                var row = sheet.GetOrCreateRow(nameManger.RowIndex - 1);
                var cell = row.GetOrCreateCell(nameManger.ColumnIndex);

                cell.SetCellValue(value.ToString());


            }
        }

        public void WriteListByNameManager<T>(List<T> datas, string sheetName = "Sheet1") where T : class, new()
        {
            var sheet = Workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = Workbook.GetSheetAt(0);
            }

            var objProperties = typeof(T).GetProperties();
            var (nameManagerList, startRowIndex) = GetNameManagerList(sheet.SheetName);

            for (int i = 0; i < datas.Count; i++)
            {
                var elData = datas.ElementAt(i);

                var row = sheet.GetOrCreateRow(startRowIndex + i);
                foreach (var item in nameManagerList)
                {
                    object value = null;
                    var property = objProperties.FirstOrDefault(e => e.Name == item.Name);
                    if (property == null)
                    {
                        continue;
                    }
                    else
                    {
                        value = property.GetValue(elData);
                    }

                    row.GetOrCreateCell(item.ColumnIndex).SetCellValue(value?.ToString());
                }
            }
        }


        public static object Convert(ICell cell, Type valueType)
        {

            object value = null;
            switch (cell.CellType)
            {
                case CellType.Unknown:
                    break;
                case CellType.Numeric:
                    value = cell.NumericCellValue;
                    break;
                case CellType.String:
                    value = cell.StringCellValue ?? "";
                    break;
                case CellType.Formula:
                    break;
                case CellType.Blank:
                    break;
                case CellType.Boolean:
                    value = cell.BooleanCellValue;
                    break;
                case CellType.Error:
                    break;
                default:
                    break;
            }

            valueType = valueType.GenericTypeArguments.Length > 0 ? valueType.GenericTypeArguments[0] : valueType;

            if (value != null && typeof(Enum).IsAssignableFrom(valueType))
            {
                value = EnumExtends.GetEnumByEnumDisplayName(valueType, value?.ToString());
            }
            if (value != null && typeof(bool) == valueType && bool.TryParse(value?.ToString(), out bool result) == false)
            {
                if (value?.ToString() == "是")
                {
                    result = true;
                }
                value = result;
            }
            if (value == null)
            {
                if (typeof(string).IsAssignableFrom(valueType))
                {
                    value = "";
                }
                else if (typeof(DateTime).IsAssignableFrom(valueType))
                {
                    value = default(DateTime);
                }
                else if (typeof(decimal).IsAssignableFrom(valueType))
                {
                    value = default(decimal);
                }
                else if (typeof(double).IsAssignableFrom(valueType))
                {
                    value = default(double);
                }
                else if (typeof(float).IsAssignableFrom(valueType))
                {
                    value = default(float);
                }
                else if (typeof(int).IsAssignableFrom(valueType))
                {
                    value = default(int);
                }
                else if (typeof(short).IsAssignableFrom(valueType))
                {
                    value = default(short);
                }
                return value;
            }
            return System.Convert.ChangeType(value, valueType);
        }

        public static object GetDefaultValue(Type valueType)
        {
            object value = null;
            if (typeof(string).IsAssignableFrom(valueType))
            {
                value = "";
            }
            else if (typeof(DateTime).IsAssignableFrom(valueType))
            {
                value = default(DateTime);
            }
            else if (typeof(decimal).IsAssignableFrom(valueType))
            {
                value = default(decimal);
            }
            else if (typeof(double).IsAssignableFrom(valueType))
            {
                value = default(double);
            }
            else if (typeof(float).IsAssignableFrom(valueType))
            {
                value = default(float);
            }
            else if (typeof(int).IsAssignableFrom(valueType))
            {
                value = default(int);
            }
            return value;

        }

        public void Write(string filePath)
        {
            Write(File.Open(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite));
        }

        public void Write(Stream fileStream)
        {
            Workbook.Write(fileStream, false);
            Dispose();
        }

        public object GetWorkbook()
        {
            return Workbook;
        }
    }
}

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using static System.String;

namespace System
{
    /// <summary>
    /// 把枚举值按照指定的文本显示
    /// <remarks>
    /// 一般通过枚举值的ToString()可以得到变量的文本，
    /// 但是有时候需要的到与之对应的更充分的文本，
    /// 这个类帮助达到此目的
    /// </remarks>
    /// </summary>
    /// <example>
    /// [EnumDescription("中文数字")]
    /// enum MyEnum
    /// {
    ///		[EnumDescription("数字一")]
    /// 	One = 1, 
    /// 
    ///		[EnumDescription("数字二")]
    ///		Two, 
    /// 
    ///		[EnumDescription("数字三")]
    ///		Three
    /// }
    /// EnumDescription.GetEnumText(typeof(MyEnum));
    /// EnumDescription.GetFieldText(MyEnum.Two);
    /// EnumDescription.GetFieldTexts(typeof(MyEnum)); 
    /// </example>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Enum)]
    public class EnumDescription : Attribute
    {
        private FieldInfo _fieldIno;

        /// <summary>
        /// 描述枚举值
        /// </summary>
        /// <param name="enumDisplayText">描述内容</param>
        /// <param name="enumRank">排列顺序</param>
        public EnumDescription(string enumDisplayText, int enumRank)
        {
            EnumDisplayText = enumDisplayText;
            EnumRank = enumRank;
        }

        /// <summary>
        /// 描述枚举值，默认排序为5
        /// </summary>
        /// <param name="enumDisplayText">描述内容</param>
        public EnumDescription(string enumDisplayText)
            : this(enumDisplayText, 5) { }

        public string EnumDisplayText { get; }

        public int EnumRank { get; }

        private int EnumValue => (int)_fieldIno.GetValue(null);

        public string FieldName => _fieldIno.Name;

        #region  =========================================对枚举描述属性的解释相关函数

        /// <summary>
        /// 排序类型
        /// </summary>
        public enum SortType
        {
            /// <summary>
            ///按枚举顺序默认排序
            /// </summary>
            Default,
            /// <summary>
            /// 按描述值排序
            /// </summary>
            DisplayText,
            /// <summary>
            /// 按排序熵
            /// </summary>
            Rank
        }

        private static readonly Hashtable CachedEnum = new Hashtable();


        /// <summary>
        /// 得到对枚举的描述文本
        /// </summary>
        /// <param name="enumType">枚举类型</param>
        /// <returns></returns>
        public static string GetEnumText(Type enumType)
        {
            EnumDescription[] eds = (EnumDescription[])enumType.GetCustomAttributes(typeof(EnumDescription), false);
            if (eds.Length != 1) return Empty;
            return eds[0].EnumDisplayText;
        }

        /// <summary>
        /// 获得指定枚举类型中，指定值的描述文本。
        /// </summary>
        /// <param name="enumValue">枚举值，不要作任何类型转换</param>
        /// <returns>描述字符串</returns>
        public static string GetFieldText(object enumValue)
        {
            EnumDescription[] descriptions = GetFieldTexts(enumValue.GetType(), SortType.Default);
            foreach (EnumDescription ed in descriptions)
            {
                if (ed._fieldIno.Name == enumValue.ToString()) return ed.EnumDisplayText;
            }
            return Empty;
        }


        /// <summary>
        /// 得到枚举类型定义的所有文本，按定义的顺序返回
        /// </summary>
        /// <exception cref="NotSupportedException"></exception>
        /// <param name="enumType">枚举类型</param>
        /// <returns>所有定义的文本</returns>
        public static EnumDescription[] GetFieldTexts(Type enumType)
        {
            return GetFieldTexts(enumType, SortType.Default);
        }

        /// <summary>
        /// 得到枚举类型定义的所有文本
        /// </summary>
        /// <exception cref="NotSupportedException"></exception>
        /// <param name="enumType">枚举类型</param>
        /// <param name="sortType">指定排序类型</param>
        /// <returns>所有定义的文本</returns>
        public static EnumDescription[] GetFieldTexts(Type enumType, SortType sortType)
        {
            //缓存中没有找到，通过反射获得字段的描述信息
            if (CachedEnum.Contains(enumType.FullName) == false)
            {
                FieldInfo[] fields = enumType.GetFields();
                ArrayList edAl = new ArrayList();
                foreach (FieldInfo fi in fields)
                {
                    object[] eds = fi.GetCustomAttributes(typeof(EnumDescription), false);
                    if (eds.Length != 1) continue;
                    ((EnumDescription)eds[0])._fieldIno = fi;
                    edAl.Add(eds[0]);
                }

                CachedEnum.Add(enumType.FullName, (EnumDescription[])edAl.ToArray(typeof(EnumDescription)));
            }
            var descriptions = (EnumDescription[])CachedEnum[enumType.FullName];
            if (descriptions.Length <= 0) throw new NotSupportedException("枚举类型[" + enumType.Name + "]未定义属性EnumValueDescription");

            //按指定的属性冒泡排序
            for (int m = 0; m < descriptions.Length; m++)
            {
                //默认就不排序了
                if (sortType == SortType.Default) break;

                for (int n = m; n < descriptions.Length; n++)
                {
                    EnumDescription temp;
                    bool swap = false;

                    switch (sortType)
                    {
                        case SortType.Default:
                            break;
                        case SortType.DisplayText:
                            if (CompareOrdinal(descriptions[m].EnumDisplayText, descriptions[n].EnumDisplayText) > 0) swap = true;
                            break;
                        case SortType.Rank:
                            if (descriptions[m].EnumRank > descriptions[n].EnumRank) swap = true;
                            break;
                    }

                    if (swap)
                    {
                        temp = descriptions[m];
                        descriptions[m] = descriptions[n];
                        descriptions[n] = temp;
                    }
                }
            }

            return descriptions;
        }

        #endregion
        #region 转换枚举成LIST
        public static IList<KeyValuePair<string, T>> ToList<T>()
        {
            T[] types = (T[])Enum.GetValues(typeof(T));
            return types.Select(type => new KeyValuePair<string, T>(GetFieldText(type), type)).ToList();
        }
        public static IList<KeyValuePair<string, int>> ToList(string typeName)
        {
            var type = Type.GetType(typeName, true, true);
            var types = Enum.GetValues(type);
            List<KeyValuePair<string, int>> list = new List<KeyValuePair<string, int>>();
            foreach (var item in types)
            {
                list.Add(new KeyValuePair<string, int>(GetFieldText(item), (int)item));
            }
            return list;
        }
        #endregion

        #region 枚举值转换成描述
        /// <summary>
        /// 获取枚举值的描述
        /// </summary>
        public static string GetEnumName<T>(object value)
        {
            T obj = (T)value;
            string enumName = GetFieldText(obj);
            return enumName;
        }

        #endregion

        #region 枚举描述转换成枚举值

        /// <summary>
        /// 根据枚举描述，获取枚举值的值
        /// </summary>
        public static object GetEnumValue<T>(string text)
        {
            EnumDescription[] descriptions = GetFieldTexts(typeof(T), SortType.Default);
            foreach (EnumDescription ed in descriptions)
            {
                if (ed.EnumDisplayText == text) return ed.EnumValue;
            }
            return null;
        }

        #endregion

        public static string GetEnumDescripByVal<T>(string enumString)
        {
            var result = string.Empty;
            if (!string.IsNullOrEmpty(enumString))
            {
                var e = (T)Enum.Parse(typeof(T), enumString);
                result = GetEnumName<T>(e);
                if (string.IsNullOrEmpty(result))
                    result = enumString;
            }
            return result;
        }

    }

    public static class IntExtensions
    {
        public static bool ToBool(this int value)
        {
            return value == 1;
        }
    }
}

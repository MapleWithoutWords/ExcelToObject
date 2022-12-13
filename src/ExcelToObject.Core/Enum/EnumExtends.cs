using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace System
{
    #region Enum扩展

    public static class EnumExtends
    {
        public static string GetDisplayName(this System.Enum enumItem)
        {
            return GetDisplayName(enumItem, null);
        }

        public static string GetDisplayName(this System.Enum enumItem, string candidate)
        {
            var field = enumItem.GetType().GetField(enumItem.ToString());
            if (field == null)
            {
                if (candidate != null) return candidate;
                return enumItem.ToString();
            }

            var attributes = field.GetCustomAttributes(typeof(EnumDescription), false);
            foreach (var attribute in attributes)
            {
                if (attribute is EnumDescription)
                {
                    return ((EnumDescription)attribute).EnumDisplayText;
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// 通过EnumDisplayName得到枚举
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="enumDisplayName"></param>
        /// <returns></returns>
        public static T GetEnumByEnumDisplayName<T>(string enumDisplayName)
        {
            Type _type = typeof(T);
            foreach (var field in _type.GetFields())
            {
                var cus = field.CustomAttributes.FirstOrDefault();
                if (cus != null && cus.ConstructorArguments[0].Value.ToString() == enumDisplayName)
                {
                    return (T)field.GetValue(null);
                }
                else
                {
                    if (field.Name == enumDisplayName)
                        return (T)field.GetValue(null);
                }
            }
            throw new ArgumentException($"{enumDisplayName} 未能找到对应的枚举.", "DisplayName");
        }
        /// <summary>
        /// 通过EnumDisplayName得到枚举
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="enumDisplayName"></param>
        /// <returns></returns>
        public static object GetEnumByEnumDisplayName(Type enumType, string enumDisplayName)
        {
            Type _type = enumType;
            foreach (var field in _type.GetFields())
            {
                var cus = field.CustomAttributes.FirstOrDefault();
                if (cus != null && cus.ConstructorArguments[0].Value.ToString() == enumDisplayName)
                {
                    return field.GetValue(null);
                }
                else
                {
                    if (field.Name == enumDisplayName)
                        return field.GetValue(null);
                }
            }
            return default;
        }


        public static T GetEnumByValue<T>(string enumValue) where T:struct
        {
            if(Enum.TryParse(enumValue, out T t))
            {
                return t;
            }

            throw new ArgumentException($"{enumValue} 未能找到对应的枚举.{typeof(T).FullName.ToString()}");
        }



        public static List<EnumberEntity> EnumToList<T>()
        {
            List<EnumberEntity> list = new List<EnumberEntity>();

            foreach (var e in Enum.GetValues(typeof(T)))
            {
                EnumberEntity m = new EnumberEntity();
                var attributes = e.GetType().GetField(e.ToString()).GetCustomAttributes(typeof(EnumDescription), true);
                foreach (var attribute in attributes)
                {
                    if (attribute is EnumDescription)
                    {
                        m.Desction = ((EnumDescription)attribute).EnumDisplayText;
                    }
                }
                m.EnumValue = Convert.ToInt32(e);
                m.EnumName = e.ToString();
                list.Add(m);
            }
            return list;
        }

        public class EnumberEntity
        {
            /// <summary>  
            /// 枚举的描述  
            /// </summary>  
            public string Desction { set; get; }

            /// <summary>  
            /// 枚举名称  
            /// </summary>  
            public string EnumName { set; get; }

            /// <summary>  
            /// 枚举对象的值  
            /// </summary>  
            public int EnumValue { set; get; }
        }
    }

    [AttributeUsage(AttributeTargets.Enum | AttributeTargets.Field, AllowMultiple = true)]
    public class EnumFilterAttribute : Attribute
    {
        public EnumFilterAttribute(string name)
        {
            Name = name;
        }

        public string Name { get; private set; }
    }

    #endregion Enum扩展
}

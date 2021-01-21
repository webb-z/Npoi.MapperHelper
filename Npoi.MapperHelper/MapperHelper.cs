using System;
using Microsoft.AspNetCore.Mvc;
using Npoi.Mapper;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.AspNetCore.Http;

namespace Npoi.MapperHelper
{
    public class MapperHelper
    {
        /// <summary>
        /// Excel输出  属性上使用特性[Description]才能够反射名称不然返回实体名称
        /// 单页码
        /// </summary>
        /// <typeparam name="T">返回Dto类型</typeparam>
        /// <param name="obj">Dto</param>
        /// <param name="fileName">FileName</param>
        /// <param name="contentType">ContentType</param>
        /// <returns></returns>
        public FileResult Export<T>(IEnumerable<T> obj, string fileName, string contentType, string sheetName)
        {
            var mapper = new Npoi.Mapper.Mapper();
            foreach (var propertyInfo in typeof(T).GetProperties())
            {
                if (propertyInfo.CustomAttributes.ToArray().Count() > 0 && propertyInfo.CustomAttributes.ToArray()[0].ConstructorArguments.Count > 0)
                {
                    var value = propertyInfo.CustomAttributes.ToArray()[0].ConstructorArguments[0].Value;
                    mapper.Map<T>(value.ToString(), propertyInfo.Name);
                }
                else
                    mapper.Map<T>(propertyInfo.Name, propertyInfo.Name);
            }
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            mapper.Save(stream, obj, sheetName, overwrite: true, xlsx: true);
            return new FileContentResult(stream.ToArray(), contentType) { FileDownloadName = fileName };
        }

        /// <summary>
        /// 文件流导入
        /// </summary>
        /// <typeparam name="T">导出Dto</typeparam>
        /// <param name="file">文件流</param>
        /// <returns></returns>
        public List<T> Import<T>(IFormFile file)
        {
            var mapper = new Npoi.Mapper.Mapper(file.OpenReadStream());
            var work = mapper.Workbook;
            //字典文件转换为Mapper
            Dictionary<string, string> dic = new Dictionary<string, string>();
            foreach (var propertyInfo in typeof(T).GetProperties())
            {
                if (propertyInfo.CustomAttributes.ToArray().Count() > 0 && propertyInfo.CustomAttributes.ToArray()[0].ConstructorArguments.Count > 0)
                {
                    mapper.Map<T>(propertyInfo.CustomAttributes.ToArray()[0].ConstructorArguments[0].Value.ToString(), propertyInfo.Name);
                    dic.Add(propertyInfo.CustomAttributes.ToArray()[0].ConstructorArguments[0].Value.ToString(), propertyInfo.Name);
                }
                else
                {
                    mapper.Map<T>(propertyInfo.Name, propertyInfo.Name);
                    dic.Add(propertyInfo.Name, propertyInfo.Name);
                }
            }

            var valueList = mapper.Take<dynamic>(work.GetSheetName(0)).Select(d => d.Value).ToList();
            List<T> list = new List<T>();
            foreach (var value in valueList)
            {
                T t = Activator.CreateInstance<T>();
                foreach (System.Reflection.PropertyInfo p in value.GetType().GetProperties())
                {
                    //字典内存的值为List 的属性名
                    var dicName = dic.Where(d => d.Key == p.Name).Select(d => d.Value).FirstOrDefault();
                    System.Reflection.PropertyInfo pro = typeof(T).GetProperties().FirstOrDefault(d => d.Name == dicName);
                    //pro.PropertyType.
                    pro.SetValue(t, ValueConversion(p.GetValue(value, null), pro));
                }
                list.Add(t);
            }

            return list;
        }


        /// <summary>
        /// 返回值类型
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="Property">属性</param>
        /// <returns></returns>
        private static dynamic ValueConversion(dynamic value, PropertyInfo Property)
        {
            if (Property.PropertyType.Equals(typeof(string)))
            {
                return value.ToString();
            }
            else if (Property.PropertyType.Equals(typeof(int)))
            {
                return (int)value;
            }
            else if (Property.PropertyType.Equals(typeof(decimal)))
            {
                return (decimal)value;
            }
            else if (Property.PropertyType.Equals(typeof(DateTime)))
            {
                return (DateTime)value;
            }
            else if (Property.PropertyType.Equals(typeof(double)))
            {
                return (double)value;
            }
            else if (Property.PropertyType.Equals(typeof(bool)))
            {
                return (bool)value;
            }
            else
                return null;
            return null;
        }
    }
}

using Microsoft.AspNetCore.Mvc;
using Npoi.Mapper;
using System.Collections.Generic;
using System.Linq;

namespace Npoi.MapperHelper
{
    public class MapperHelper
    {
        /// <summary>
        /// Excel输出  属性上使用特性[Description]才能够反射名称不然返回实体名称
        /// </summary>
        /// <typeparam name="T">返回Dto类型</typeparam>
        /// <param name="obj">Dto</param>
        /// <param name="fileName">FileName</param>
        /// <param name="contentType">ContentType</param>
        /// <returns></returns>
        public FileResult Export<T>(IEnumerable<T> obj, string fileName, string contentType)
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
            mapper.Save(stream, obj, "sheet1", overwrite: true, xlsx: true);
            return new FileContentResult(stream.ToArray(), contentType) { FileDownloadName = fileName };
        }
    }
}

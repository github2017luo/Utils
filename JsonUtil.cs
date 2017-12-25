using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MonitorAndon.Common
{
    public class JsonUtil<T> where T:new()
    {
        //install-package Newtonsoft.Json 
     
        /// <summary>
        /// 将Json串序列化成对应的实体(在线Json生成C#实体)
        /// </summary>
        /// <param name="JsonStr"></param>
        /// <returns></returns>
        public static T JsonStrToEntity(string JsonStr)
        {
            T entity = new T();
            entity = JsonConvert.DeserializeAnonymousType(JsonStr, new T());
            return entity;
        }
        /// <summary>
        /// 将实体序列化成Json串
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public static string EntityToJsonStr(T entity)
        {
            string temp = string.Empty;
            temp = JsonConvert.SerializeObject(entity);
            return temp;
        }
    }
}

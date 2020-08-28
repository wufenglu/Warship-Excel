using Warship.Attribute.Attributes;
using Warship.Excel.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Import.Helper
{
    /// <summary>
    /// 级联缓存
    /// </summary>
    public class CascadeCache
    {
        /// <summary>
        /// 子级实体集合
        /// </summary>
        public static Dictionary<string, IEnumerable> ChildSheetEntityList { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        static CascadeCache()
        {
            ChildSheetEntityList = new Dictionary<string, IEnumerable>();
        }
    }

    /// <summary>
    /// 级联处理
    /// </summary>
    public class Cascade<MasterT, SlaveT> 
    {
        /// <summary>
        /// 设置实体属性的值
        /// </summary>
        /// <param name="sheetAttribute"></param>
        /// <param name="entity"></param>
        /// <param name="prop"></param>
        /// <param name="childSheetEntityList"></param>
        public void SetEntityPropertyValues(SheetAttribute sheetAttribute, MasterT entity, PropertyInfo prop, List<SlaveT> childSheetEntityList)
        {
            List<SlaveT> results = new List<SlaveT>();
            //获取主实体的属性值
            string masterValue = entity.GetType().GetProperty(sheetAttribute.MasterEntityProperty).GetValue(entity, null)?.ToString();
            foreach (var item in childSheetEntityList)
            {
                //获取从属性的属性值
                string slaveValue = item.GetType().GetProperty(sheetAttribute.SlaveEntityProperty).GetValue(item, null).ToString();

                //如果两者相等说明一致，则向当前属性上赋值
                if (masterValue == slaveValue)
                {
                    results.Add(item);
                }
            }
            prop.SetValue(entity, results);
        }
    }
}

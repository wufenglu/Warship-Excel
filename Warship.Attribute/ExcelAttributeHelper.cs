using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Warship.Attribute.Attributes;
using Warship.Attribute.Model;
using Warship.Attribute.Validations;

namespace Warship.Attribute
{
    /// <summary>
    /// Excel特性帮助类
    /// </summary>
    public class ExcelAttributeHelper<TEntity> where TEntity : class, new()
    {
        /// <summary>
        /// 获取头部集合
        /// </summary>
        /// <returns></returns>
        public static List<ExcelHeadDTO> GetHeads()
        {
            //结果集
            List<ExcelHeadDTO> result = new List<ExcelHeadDTO>();

            PropertyInfo[] propertyInfos = typeof(TEntity).GetProperties();
            //遍历属性
            foreach (PropertyInfo info in propertyInfos)
            {
                //获取头部特性
                ExcelHeadAttribute attribute = info.GetCustomAttribute<ExcelHeadAttribute>();
                if (attribute != null)
                {
                    //头部实体赋值
                    ExcelHeadDTO dto = new ExcelHeadDTO();
                    dto.PropertyName = info.Name;
                    dto.HeadName = attribute.HeadName;
                    dto.IsSetHeadColor = attribute.IsSetHeadColor;                    
                    dto.IsLocked = attribute.IsLocked;
                    dto.SortValue = attribute.SortValue;
                    dto.ColumnType = attribute.ColumnType;
                    dto.IsHiddenColumn = attribute.IsHiddenColumn;
                    dto.ColumnWidth = attribute.ColumnWidth;
                    dto.Format = attribute.Format;
                    dto.BackgroundColor = attribute.BackgroundColor;
                    dto.HeaderBackgroundColor = attribute.HeaderBackgroundColor;
                    //获取属性
                    RequiredAttribute requiredAttribute = info.GetCustomAttribute<RequiredAttribute>();
                    List<string> disableAttrs = new RequiredAttributeValidation<TEntity>().GetDisableAttributes();
                    //启用必填属性
                    if (requiredAttribute != null && (disableAttrs == null || disableAttrs.Contains(info.Name) == false))
                    {
                        dto.IsValidationHead = true;
                    }
                    else
                    {
                        dto.IsValidationHead = false;
                    }

                    //启用必填属性
                    Dictionary<string, BaseAttribute> enableAttrs = new RequiredAttributeValidation<TEntity>().GetEnableAttributes();
                    if (enableAttrs != null && enableAttrs.Keys.Contains(info.Name) == true)
                    {
                        dto.IsValidationHead = true;
                    }

                    //添加至结果集
                    result.Add(dto);
                }
     
            }

            int columnIndex = 0;
            //返回结果集
            result = result.OrderBy(it => it.SortValue).ToList();
            result.ForEach(it => { it.ColumnIndex = columnIndex++; });
            return result;
        }

        /// <summary>
        /// 获取头部对象
        /// </summary>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static ExcelHeadDTO GetHead(string propertyName)
        {
            //获取属性对应的头部实体
            List<ExcelHeadDTO> list = GetHeads();
            return list.Where(w => w.PropertyName == propertyName).First();
        }
    }
}

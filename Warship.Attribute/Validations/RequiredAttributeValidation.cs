using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Warship.Attribute.Attributes;
using Warship.Attribute.Model;
using Warship.Utility;

namespace Warship.Attribute.Validations
{
    /// <summary>
    /// 必填校验
    /// </summary>
    public class RequiredAttributeValidation<TEntity> : BaseAttributeValication<TEntity>
        where TEntity : class
    {
        /// <summary>
        /// 属性键
        /// </summary>
        /// <returns></returns>
        protected override string GetAttributeKey()
        {
            return "Required";
        }

        /// <summary>
        /// 验证
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public override List<ValidationResult> Validation(TEntity entity)
        {
            List<ValidationResult> result = new List<ValidationResult>();
            List<string> disableAttributes = GetDisableAttributes();
            Dictionary<string, BaseAttribute> enableAttributes = GetEnableAttributes();

            //遍历属性
            foreach (PropertyInfo prop in PropertyHelper.GetPropertys<TEntity>())
            {
                //获取属性
                RequiredAttribute requiredAttribute = prop.GetCustomAttribute<RequiredAttribute>();
                if (requiredAttribute == null && enableAttributes != null)
                {
                    //实体没有标识必填特性，且没有设置启用必填特性，则跳出
                    if (enableAttributes.Keys.Contains(prop.Name) == true)
                    {
                        requiredAttribute = enableAttributes[prop.Name] as RequiredAttribute;
                    }
                }

                //为空则跳出
                if (requiredAttribute == null)
                {
                    continue;
                }

                //如果是禁用特性则跳出
                if (disableAttributes != null && disableAttributes.Contains(prop.Name))
                {
                    continue;
                }

                //获取属性值
                object value = prop.GetValueExt(entity);// entity.GetPropertyValue(prop.Name);

                //必填校验
                if (value == null || string.IsNullOrEmpty(value.ToString()))
                {
                    result.Add(new ValidationResult()
                    {
                        PropertyName = prop.Name,
                        ErrorMessage = requiredAttribute.ErrorMessage
                    });
                }
            }
            
            return result;
        }
    }
}

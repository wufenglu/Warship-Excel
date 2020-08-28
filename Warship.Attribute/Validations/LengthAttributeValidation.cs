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
    /// 长度校验
    /// </summary>
    public class LengthAttributeValidation<TEntity> : BaseAttributeValication<TEntity>
        where TEntity : class
    {
        /// <summary>
        /// 属性键
        /// </summary>
        /// <returns></returns>
        protected override string GetAttributeKey()
        {
            return "Length";
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
                LengthAttribute lengthAttribute = prop.GetCustomAttribute<LengthAttribute>();
                if (lengthAttribute == null && enableAttributes != null)
                {
                    //实体没有标识必填特性，且没有设置启用必填特性，则跳出
                    if (enableAttributes.Keys.Contains(prop.Name) == true)
                    {
                        lengthAttribute = enableAttributes[prop.Name] as LengthAttribute;
                    }
                }

                //为空则跳出
                if (lengthAttribute == null)
                {
                    continue;
                }

                //如果是禁用特性则跳出
                if (disableAttributes != null && disableAttributes.Contains(prop.Name))
                {
                    continue;
                }


                //获取属性值
                object value = prop.GetValueExt(entity);

                //为空则跳出
                if (value == null)
                {
                    continue;
                }

                //长度校验
                if (lengthAttribute.Length < value.ToString().Length)
                {
                    result.Add(new ValidationResult()
                    {
                        PropertyName = prop.Name,
                        ErrorMessage = lengthAttribute.ErrorMessage
                    });
                }
            }

            return result;
        }
    }
}

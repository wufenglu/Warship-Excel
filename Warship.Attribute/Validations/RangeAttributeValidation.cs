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
    public class RangeAttributeValidation<TEntity> : BaseAttributeValication<TEntity>
        where TEntity : class
    {
        /// <summary>
        /// 属性键
        /// </summary>
        /// <returns></returns>
        protected override string GetAttributeKey()
        {
            return "Range";
        }

        /// <summary>
        /// 验证
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public override List<ValidationResult> Validation(TEntity entity)
        {
            //变量定义
            List<ValidationResult> result = new List<ValidationResult>();
            List<string> disableAttributes = GetDisableAttributes();
            Dictionary<string, BaseAttribute> enableAttributes = GetEnableAttributes();

            //遍历属性
            foreach (PropertyInfo prop in PropertyHelper.GetPropertys<TEntity>())
            {
                //获取特性
                RangeAttribute rangeAttribute = prop.GetCustomAttribute<RangeAttribute>();
                if (rangeAttribute == null && enableAttributes != null)
                {
                    //实体没有标识必填特性，且没有设置启用必填特性，则跳出
                    if (enableAttributes.Keys.Contains(prop.Name) == true)
                    {
                        rangeAttribute = enableAttributes[prop.Name] as RangeAttribute;
                    }
                }

                //为空则跳出
                if (rangeAttribute == null)
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

                //如果是可为空类型，且值为空则跳出
                if (prop.PropertyType.IsGenericType
                    && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>)
                    && value == null)
                {
                    continue;
                }

                //判断值是否为空
                if (value == null)
                {
                    //添加异常
                    result.Add(new ValidationResult()
                    {
                        PropertyName = prop.Name,
                        ErrorMessage = rangeAttribute.ErrorMessage
                    });
                    continue;
                }

                //是否有错误
                bool isError = CheckIsError(rangeAttribute, value);
                //是否有异常
                if (isError)
                {
                    result.Add(new ValidationResult()
                    {
                        PropertyName = prop.Name,
                        ErrorMessage = rangeAttribute.ErrorMessage
                    });
                }
            }
            return result;
        }

        /// <summary>
        /// 检查范围
        /// </summary>
        /// <param name="rangeAttribute"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool CheckIsError(RangeAttribute rangeAttribute, object value)
        {
            try
            {
                //比较大小
                decimal decimalMin = Convert.ToDecimal(rangeAttribute.Minimum);
                decimal decimalMax = Convert.ToDecimal(rangeAttribute.Maximum);
                decimal decimalPropValue = Convert.ToDecimal(value);
                if (decimalPropValue < decimalMin || decimalPropValue > decimalMax)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                string err = ex.StackTrace;
                return true;
            }
            return false;
        }
    }
}

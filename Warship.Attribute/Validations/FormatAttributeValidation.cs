using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Warship.Attribute.Attributes;
using Warship.Attribute.Enum;
using Warship.Attribute.Model;
using Warship.Utility;

namespace Warship.Attribute.Validations
{
    /// <summary>
    /// 格式校验属性
    /// </summary>
    /// <typeparam name="TEntity"></typeparam>
    public class FormatAttributeValidation<TEntity> : BaseAttributeValication<TEntity>
        where TEntity : class
    {
        /// <summary>
        /// 属性键
        /// </summary>
        /// <returns></returns>
        protected override string GetAttributeKey()
        {
            return "Format";
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
                FormatAttribute formatAttribute = prop.GetCustomAttribute<FormatAttribute>();
                if (formatAttribute == null && enableAttributes != null)
                {
                    //实体没有标识必填特性，且没有设置启用必填特性，则跳出
                    if (enableAttributes.Keys.Contains(prop.Name) == true)
                    {
                        formatAttribute = enableAttributes[prop.Name] as FormatAttribute;
                    }
                }

                //为空则跳出
                if (formatAttribute == null)
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

                if (value == null)
                {
                    result.Add(new ValidationResult()
                    {
                        PropertyName = prop.Name,
                        ErrorMessage = formatAttribute.ErrorMessage
                    });
                    continue;
                }

                //正则表达式
                string regex = null;
                switch (formatAttribute.FormatEnum)
                {
                    case FormatEnum.Email:
                        regex = "^\\s*([A-Za-z0-9_-]+(\\.\\w+)*@(\\w+\\.)+\\w{2,5})\\s*$";
                        break;
                    case FormatEnum.Phone:
                        regex = @"^(\d{3,4}-)?\d{6,8}$";
                        break;
                    case FormatEnum.MobilePhone:
                        regex = @"^[1]+[3,5,8]+\d{9}";
                        break;
                    case FormatEnum.ID:
                        regex = @"(^\d{18}$)|(^\d{15}$)";
                        break;
                }

                //如果用户配置了正则表达式，以用户配置为准
                if (string.IsNullOrEmpty(formatAttribute.Regex) == false)
                {
                    regex = formatAttribute.Regex;
                }

                //如果正则为空则跳出
                if (regex == null)
                {
                    continue;
                }

                Regex r = new Regex(regex);
                //如果匹配不上则添加异常
                if (r.IsMatch(value.ToString()) == false)
                {
                    result.Add(new ValidationResult()
                    {
                        PropertyName = prop.Name,
                        ErrorMessage = formatAttribute.ErrorMessage
                    });
                }
            }
            return result;
        }

        /// <summary>
        /// 格式检查
        /// </summary>
        /// <param name="formatAttribute"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool CheckIsError(FormatAttribute formatAttribute,object value) {
            //正则表达式
            string regex = null;
            switch (formatAttribute.FormatEnum)
            {
                case FormatEnum.Email:
                    regex = "^\\s*([A-Za-z0-9_-]+(\\.\\w+)*@(\\w+\\.)+\\w{2,5})\\s*$";
                    break;
                case FormatEnum.Phone:
                    regex = @"^(\d{3,4}-)?\d{6,8}$";
                    break;
                case FormatEnum.MobilePhone:
                    regex = @"^[1]+[3,5,8]+\d{9}";
                    break;
                case FormatEnum.ID:
                    regex = @"(^\d{18}$)|(^\d{15}$)";
                    break;
            }

            //如果用户配置了正则表达式，以用户配置为准
            if (string.IsNullOrEmpty(formatAttribute.Regex) == false)
            {
                regex = formatAttribute.Regex;
            }

            //如果正则为空则跳出
            if (regex == null)
            {
                return false;
            }

            Regex r = new Regex(regex);
            //如果匹配不上则添加异常
            if (r.IsMatch(value.ToString()) == false)
            {
                return true;
            }

            return false;
        }
    }
}

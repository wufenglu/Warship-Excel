using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.ComponentModel;
using Warship.Attribute.Validations;
using Warship.Attribute.Enum;

namespace Warship.Attribute
{
    /// <summary>
    /// 属性工厂
    /// </summary>
    public static class AttributeFactory<TEntity> where TEntity:class
    {
        /// <summary>
        /// 验证
        /// </summary>
        /// <param name="validationType"></param>
        /// <returns></returns>
        public static IAttributeValidation<TEntity> GetValication(ValidationTypeEnum validationType)
        {
            switch (validationType)
            {
                case ValidationTypeEnum.Required:
                    return new RequiredAttributeValidation<TEntity>();
                case ValidationTypeEnum.Length:
                    return new LengthAttributeValidation<TEntity>();
                case ValidationTypeEnum.Range:
                    return new RangeAttributeValidation<TEntity>();
                case ValidationTypeEnum.Format:
                    return new FormatAttributeValidation<TEntity>();
            }
            return null;
        }
    }
}

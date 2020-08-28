using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Attribute.Attributes;
using Warship.Attribute.Model;

namespace Warship.Attribute.Validations
{
    /// <summary>
    /// 基础校验接口
    /// </summary>
    /// <typeparam name="TEntity"></typeparam>
    public interface IAttributeValidation<TEntity> where TEntity : class
    {
        /// <summary>
        /// 验证
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        List<ValidationResult> Validation(TEntity entity);

        /// <summary>
        /// 禁用属性
        /// </summary>
        /// <param name="attributeNames"></param>
        void DisableAttributes(List<string> attributeNames);

        /// <summary>
        /// 启用属性
        /// </summary>
        /// <param name="attributeNames"></param>
        void EnableAttributes(Dictionary<string, BaseAttribute> attributeNames);
    }
}

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
    /// 基类
    /// </summary>
    public abstract class BaseAttributeValication<TEntity> : IAttributeValidation<TEntity> where TEntity : class
    {
        /// <summary>
        /// 禁用的特性：扩展支撑
        /// </summary>
        public static Dictionary<string, List<string>> EntityDisableAttributes { get; set; }

        /// <summary>
        /// 启用特性：扩展支撑
        /// </summary>
        public static Dictionary<string, Dictionary<string, BaseAttribute>> EntityEnableAttributes { get; set; }
        
        /// <summary>
        /// 构造函数
        /// </summary>
        public BaseAttributeValication()
        {
            if (EntityDisableAttributes == null)
            {
                EntityDisableAttributes = new Dictionary<string, List<string>>();
            }
            if (EntityEnableAttributes == null)
            {
                EntityEnableAttributes = new Dictionary<string, Dictionary<string, BaseAttribute>>();
            }
        }

        /// <summary>
        /// 获取key
        /// </summary>
        /// <returns></returns>
        protected abstract string GetAttributeKey();

        /// <summary>
        /// 获取实体属性键
        /// </summary>
        /// <returns></returns>
        protected string GetEntityAttributeKey()
        {
            Type entity = typeof(TEntity);
            return GetAttributeKey() + "_" + entity.FullName;
        }

        #region 禁用

        /// <summary>
        /// 获取禁用特性
        /// </summary>
        /// <returns></returns>
        public List<string> GetDisableAttributes()
        {
            string entityKey = GetEntityAttributeKey();
            if (EntityDisableAttributes.Keys.Contains(entityKey))
            {
                return EntityDisableAttributes[entityKey];
            }
            return null;
        }

        /// <summary>
        /// 禁用特性
        /// </summary>
        /// <param name="attributeNames"></param>
        public void DisableAttributes(List<string> attributeNames)
        {
            string entityKey = GetEntityAttributeKey();
            if (EntityDisableAttributes.Keys.Contains(entityKey))
            {
                EntityDisableAttributes[entityKey] = attributeNames;
            }
            else {
                EntityDisableAttributes.Add(entityKey, attributeNames);
            }                
        }

        #endregion

        #region 启用

        /// <summary>
        /// 获取启用特性
        /// </summary>
        /// <returns></returns>
        public Dictionary<string, BaseAttribute> GetEnableAttributes()
        {
            string entityKey = GetEntityAttributeKey();
            if (EntityEnableAttributes.Keys.Contains(entityKey))
            {
                return EntityEnableAttributes[entityKey];
            }
            return null;
        }

        /// <summary>
        /// 启用特性:此处不考虑线程安全，一个DTO只会在一个场景使用
        /// </summary>
        /// <param name="attributeNames"></param>
        public void EnableAttributes(Dictionary<string, BaseAttribute> attributeNames)
        {
            string entityKey = GetEntityAttributeKey();

            //判断键是否存在
            if (EntityEnableAttributes.Keys.Contains(entityKey) == false)
            {
                EntityEnableAttributes.Add(entityKey, attributeNames);
                return;
            }

            //添加属性特性
            Dictionary<string, BaseAttribute> entityAttributeDict = EntityEnableAttributes[entityKey];
            string[] keyArr = attributeNames.Keys.ToArray();
            foreach (var key in keyArr) {
                entityAttributeDict[key] = attributeNames[key];
            }
        }

        #endregion

        /// <summary>
        /// 验证
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public abstract List<ValidationResult> Validation(TEntity entity);
    }
}

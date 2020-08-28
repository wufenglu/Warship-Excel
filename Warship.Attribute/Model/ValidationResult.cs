using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.ComponentModel;

namespace Warship.Attribute.Model
{
    /// <summary>
    /// 验证结果
    /// </summary>
    public class ValidationResult {
        /// <summary>
        /// 属性名称
        /// </summary>
        public string PropertyName { get; set; }
        /// <summary>
        /// 提示内容
        /// </summary>
        public string ErrorMessage { get; set; }
    }
}

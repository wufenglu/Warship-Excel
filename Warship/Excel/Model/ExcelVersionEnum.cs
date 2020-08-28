using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Model
{
    /// <summary>
    /// Excel版本
    /// </summary>
    public enum ExcelVersionEnum
    {
        /// <summary>
        /// 无版本（异常）
        /// </summary>
        No,
        /// <summary>
        /// excel2003（正常）
        /// </summary>
        V2003,
        /// <summary>
        /// excel2007（正常）
        /// </summary>
        V2007
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Excel.Model;

namespace Warship.Excel.Import
{
    /// <summary>
    /// 导入接口
    /// </summary>
    public interface IImport<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 头部验证后
        /// </summary>
        /// <param name="ExcelGlobalDTO"></param>
        void ValidationHeaderAfter(ExcelGlobalDTO<TEntity> ExcelGlobalDTO);

        /// <summary>
        /// 头部验证后
        /// </summary>
        /// <param name="ExcelGlobalDTO"></param>
        void ValidationValueAfter(ExcelGlobalDTO<TEntity> ExcelGlobalDTO);
    }
}

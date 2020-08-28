using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Excel.Model;

namespace Warship.Excel.Export
{
    /// <summary>
    /// 导出接口
    /// </summary>
    public interface IExport<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 导出扩展
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        void ExcelHandleAfter(ExcelGlobalDTO<TEntity> excelGlobalDTO);

        /// <summary>
        /// 导出扩展
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        void ExcelHandleBefore(ExcelGlobalDTO<TEntity> excelGlobalDTO);
    }
}

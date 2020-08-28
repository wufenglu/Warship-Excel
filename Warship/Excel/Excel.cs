using Warship.Excel.Import;
using Warship.Excel.Export;
using Warship.Excel.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel
{
    /// <summary>
    /// Excel对象
    /// </summary>
    public class Excel<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 导入对象
        /// </summary>
        public Import<TEntity> Import { get; set; }

        /// <summary>
        /// 导入对象
        /// </summary>
        public ImportByConfig<TEntity> ImportByConfig { get; set; }

        /// <summary>
        /// 导出对象
        /// </summary>
        public Export<TEntity> Export { get; set; }

        /// <summary>
        /// 构造函数初始化
        /// </summary>
        public Excel()
        {
            Import = new Import<TEntity>();
            ImportByConfig = new ImportByConfig<TEntity>();
            Export = new Export<TEntity>();
        }
    }
}

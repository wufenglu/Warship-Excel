using Warship.Attribute.Enum;
using Warship.Attribute.Model;
using Warship.Excel.Model.Column;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Model
{
    /// <summary>
    /// 基类实体
    /// </summary>
    public class ExcelRowModel
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public ExcelRowModel()
        {
            OtherColumns = new List<ColumnModel>();
            ColumnErrorMessage = new List<Model.ColumnErrorMessage>();
            ColumnOptions = new Dictionary<string, List<string>>();
        }

        /// <summary>
        /// 行号
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// 列错误信息
        /// </summary>
        public List<ColumnErrorMessage> ColumnErrorMessage { get; set; }

        /// <summary>
        /// 列选项值:键为列名，值为集合（例如：Type，工程合同、采购合同）
        /// </summary>
        public Dictionary<string, List<string>> ColumnOptions { get; set; }

        /// <summary>
        /// 是否删除行
        /// </summary>
        public bool IsDeleteRow { get; set; }

        /// <summary>
        /// 行样式设置
        /// </summary>
        public RowStyleSet RowStyleSet { get; set; }

        /// <summary>
        /// 其他属性
        /// </summary>
        public List<ColumnModel> OtherColumns { get; set; }

        /// <summary>
        /// 动态列：可动态添加列，适用于二开扩展增加列
        /// </summary>
        public static Dictionary<string, List<ColumnModel>> DynamicColumns { get; set; }

        /// <summary>
        /// 获取动态列
        /// </summary>
        public List<ColumnModel> GetDynamicColumns()
        {
            //键
            string key = this.GetType().FullName;
            if (DynamicColumns != null && DynamicColumns.Keys.Contains(key))
            {
                return DynamicColumns[key];
            }

            return null;
        }

        /// <summary>
        /// 设置动态列
        /// </summary>
        /// <param name="columns"></param>
        public void SetDynamicColumns(List<ColumnModel> columns)
        {
            //为空退出
            if (columns == null)
            {
                return;
            }

            //动态列为空则实例化
            if (DynamicColumns == null)
            {
                DynamicColumns = new Dictionary<string, List<ColumnModel>>();
            }

            //键
            string key = this.GetType().FullName;

            //不存在则添加
            if (DynamicColumns.Keys.Contains(key) == false)
            {
                DynamicColumns.Add(key, columns);
            }
            else
            {
                DynamicColumns[key] = columns;
            }
        }
    }
}

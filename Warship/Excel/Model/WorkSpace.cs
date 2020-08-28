using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Model
{
    /// <summary>
    /// 工作簿区间
    /// </summary>
    public class WorkSpace
    {
        /// <summary>
        /// 开始行
        /// </summary>
        public int? MinRow { set; get; }
        /// <summary>
        /// 结束行
        /// </summary>
        public int? MaxRow { set; get; }
        /// <summary>
        /// 开始列
        /// </summary>
        public int? MinCol { set; get; }
        /// <summary>
        /// 结束列
        /// </summary>
        public int? MaxCol { set; get; }
        /// <summary>
        /// 是否绝对位置
        /// </summary>
        public bool OnlyInternal { set; get; }
    }
}

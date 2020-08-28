using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Model.Column
{
    /// <summary>
    /// execl每列的附件集合
    /// </summary>
    public class ColumnFile
    {
        /// <summary>
        /// 最小开始行
        /// </summary>
        public int MinRow { get; set; }
        /// <summary>
        /// 最大开始行
        /// </summary>
        public int MaxRow { get; set; }
        /// <summary>
        /// 最小开始列
        /// </summary>
        public int MinCol { get; set; }
        /// <summary>
        /// 最大开始列
        /// </summary>
        public int MaxCol { get; set; }

        /// <summary>
        /// 文件全名
        /// </summary>
        public string FileName { set; get; }
        /// <summary>
        /// 文件扩展名称
        /// </summary>
        public string ExtensionName { set; get; }
        /// <summary>
        /// 附件文件流
        /// </summary>
        public byte[] FileBytes { set; get; }
        /// <summary>
        /// 响应类型
        /// </summary>
        public string MimeType { set; get; }
        /// <summary>
        /// 文件顺序
        /// </summary>
        public int FileIndex { set; get; }

    }
}

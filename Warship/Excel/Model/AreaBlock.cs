using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Attribute;

namespace Warship.Excel.Model
{
    /// <summary>
    /// Execl区块
    /// </summary>
    public class AreaBlock
    {
        /// <summary>
        /// 开始行
        /// </summary>
        public int StartRowIndex { set; get; }
        /// <summary>
        /// 结束行
        /// </summary>
        public int EndRowIndex { get; set; }
        /// <summary>
        /// 开始列
        /// </summary>
        public int StartColumnIndex { get; set; }
        /// <summary>
        /// 结束列
        /// </summary>
        public int EndColumnIndex { set; get; }
        /// <summary>
        /// 高度:单行文字的高度为256，三行文字高度为：256*3
        /// </summary>
        public short? Height { set; get; }
        /// <summary>
        /// 内容
        /// </summary>
        public string Content { set; get; }

        /// <summary>
        /// 多语言Key
        /// </summary>
        public string LanguageKey { get; set; }

        /// <summary>
        /// 多语言调用程序集
        /// </summary>
        public string LanguageAssembly
        {
            set
            {
                if (string.IsNullOrEmpty(LanguageKey) == false)
                {
                    Content = LanguageHelper.GetLanguageValue(LanguageKey, value);
                }
            }
        }
    }
}

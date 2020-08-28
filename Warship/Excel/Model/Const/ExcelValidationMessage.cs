namespace Warship.Excel.Model.Const
{
    /// <summary>
    /// Common 多语言
    /// </summary>
    public class ExcelValidationMessage
    {
        /// <summary>
        /// 私有变量
        /// </summary>
        private string common_Import_TempletError;

        /// <summary>
        /// 导入失败，Excel文件格式不正确，请重新导出模板！
        /// </summary>
        public string Clgyl_Common_Import_TempletError
        {
            get
            {
                return common_Import_TempletError;
            }
            set
            {
                this.common_Import_TempletError = value;
            }
        }

        /// <summary>
        /// 私有变量
        /// </summary>
        private string common_Import_NotData;
        /// <summary>
        /// 导入的Excel中不存在数据！
        /// </summary>
        public string Clgyl_Common_Import_NotData
        {
            get
            {
                return common_Import_NotData;
            }
            set
            {
                this.common_Import_NotData = value;
            }
        }

        /// <summary>
        /// 私有变量
        /// </summary>
        private string common_Import_NotExistOptions;
        /// <summary>
        /// 输入不合法, 请选择序列中的值！
        /// </summary>
        public string Clgyl_Common_Import_NotExistOptions
        {
            get
            {
                return common_Import_NotExistOptions;
            }
            set
            {
                this.common_Import_NotExistOptions = value;
            }
        }
    }
}

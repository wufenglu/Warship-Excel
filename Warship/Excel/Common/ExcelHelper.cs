using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Excel.Model;
using Warship.Excel.Export;

namespace Warship.Excel.Common
{
    /// <summary>
    /// Excel帮助
    /// </summary>
    public static class ExcelHelper
    {
        /// <summary>
        /// 跟进文件流获取工作簿
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="versionEnum">版本</param>
        public static IWorkbook GetWorkbook(Stream stream, ExcelVersionEnum versionEnum)
        {
            //2003版本
            if (versionEnum == ExcelVersionEnum.V2003)
            {
                return new HSSFWorkbook(stream);
            }

            //2007版本
            if (versionEnum == ExcelVersionEnum.V2007)
            {
                return new XSSFWorkbook(stream);
            }
            return null;
        }

        /// <summary>
        /// 创建工作簿
        /// </summary>
        /// <param name="versionEnum"></param>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook(ExcelVersionEnum versionEnum = ExcelVersionEnum.V2007)
        {
            //2003版本
            if (versionEnum == ExcelVersionEnum.V2003)
            {
                return new HSSFWorkbook();
            }

            //2007版本
            if (versionEnum == ExcelVersionEnum.V2007)
            {
                return new XSSFWorkbook();
            }
            return null;
        }

        /// <summary>
        /// 备选项导出信息
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="sheetName">sheet名称</param>
        /// <param name="datas">导出信息</param>
        public static IWorkbook OptionExport<Entity>(IWorkbook workbook, string sheetName, List<Entity> datas) where Entity : ExcelRowModel, new()
        {
            #region 备选项
            //备选项
            ExcelGlobalDTO<Entity> excelGlobalDTO = new ExcelGlobalDTO<Entity>();
            excelGlobalDTO.SetDefaultSheet();
            excelGlobalDTO.Sheets.First().SheetName = sheetName;
            excelGlobalDTO.Sheets.First().SheetEntityList = datas;
            excelGlobalDTO.Workbook = workbook;


            //创建导出对象
            Export<Entity> export = new Export<Entity>();
            export.ExecuteByData(excelGlobalDTO);

            return excelGlobalDTO.Workbook;
            #endregion
        }

        /// <summary>
        /// 跟进文件名获取版本
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static ExcelVersionEnum GetExcelVersion(string fileName)
        {
            //2007版本
            if (fileName.ToLower().IndexOf(".xlsx") > 0)
            {
                return ExcelVersionEnum.V2007;
            }

            //2003版本
            if (fileName.ToLower().IndexOf(".xls") > 0)
            {
                return ExcelVersionEnum.V2003;
            }
            return ExcelVersionEnum.No;
        }

        /// <summary>
        /// 获取列的值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string GetCellValue(ICell cell)
        {
            //为空直接返回数据
            if (cell == null)
            {
                return string.Empty;
            }
            string value = string.Empty;
            switch (cell.CellType)
            {
                case CellType.Unknown://未知类型
                    value = cell.ToString();
                    break;
                case CellType.Numeric://数字类型
                    if (HSSFDateUtil.IsCellDateFormatted(cell))
                    {
                        value = cell.DateCellValue.ToShortDateString();
                    }
                    else
                    {
                        value = cell.ToString();
                    }
                    break;
                case CellType.String://string类型
                    value = cell.StringCellValue;
                    break;
                case CellType.Formula://计算公式
                    value = string.Empty;
                    break;
                case CellType.Blank://空文本
                    value = string.Empty;
                    break;
                case CellType.Boolean://逻辑类型
                    value = cell.ToString();
                    break;
                case CellType.Error://错误内容
                    value = string.Empty;
                    break;
                default:
                    value = string.Empty;
                    break;
            }
            return value;
        }
        /// <summary>
        /// 为了兼容老版本保留两位小数，format=2的写法，将老的写法转成最新实现
        /// </summary>
        /// <param name="format"></param>
        /// <returns></returns>
        public static string GetNewFormatString(string format)
        {
            if (format.Contains("."))
            {
                return format;
            }
            int result;
            if (int.TryParse(format, out result))
            {
                //如果能转成int类型的数据就说明用户格式化参数填写的是整数，也就是老的实现方式
                return 0.ToString("f" + result);
            }
            return format;
        }
    }
}

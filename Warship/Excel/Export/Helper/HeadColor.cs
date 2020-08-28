using Warship.Excel.Model;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Export.Helper
{
    /// <summary>
    /// 列颜色
    /// </summary>
    public class HeadColor<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 设置头部颜色
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void SetHeadColor(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            foreach (var item in excelGlobalDTO.Sheets)
            {
                ISheet sheet = excelGlobalDTO.Workbook.GetSheetAt(item.SheetIndex);
                IRow row = sheet.GetRow(item.StartRowIndex.Value);

                //为空判断
                if (item.SheetHeadList == null)
                {
                    continue;
                }

                //创建头部
                foreach (var head in item.SheetHeadList)
                {
                    if (head.IsSetHeadColor == true)
                    {
                        ICell cell = row.GetCell(head.ColumnIndex);

                        IFont font = excelGlobalDTO.Workbook.CreateFont();//创建字体样式
                        font.Color = HSSFColor.Red.Index;//设置字体颜色

                        ICellStyle style = excelGlobalDTO.Workbook.CreateCellStyle();//创建单元格样
                        style.SetFont(font);
                        cell.CellStyle = style;
                    }                    
                }
            }
        }
    }
}

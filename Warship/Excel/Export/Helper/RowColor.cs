using Warship.Excel.Model;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Record;
using System.Drawing;
using Warship.Utility;

namespace Warship.Excel.Export.Helper
{
    /// <summary>
    /// 行颜色:http://www.doc88.com/p-9119769143089.html
    /// </summary>
    public class RowColor<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 行颜色
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void SetRowColor(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            foreach (var item in excelGlobalDTO.Sheets)
            {
                ISheet sheet = excelGlobalDTO.Workbook.GetSheetAt(item.SheetIndex);

                //为空判断
                if (item.SheetHeadList == null || item.SheetEntityList == null)
                {
                    continue;
                }

                //多线程处理
                MultiThreadingHelper.ForEach(item.SheetEntityList, entity =>
                {
                    if (entity.RowStyleSet == null)
                    {
                        return;
                    }
                    IRow row = sheet.GetRow(entity.RowNumber);
                    ICellStyle style = excelGlobalDTO.Workbook.CreateCellStyle();//创建单元格样

                    //创建头部
                    foreach (var head in item.SheetHeadList)
                    {
                        //列
                        var cell = row.GetCell(head.ColumnIndex);
                        if (cell == null)
                        {
                            continue;
                        }

                        //如果列有设置背景色，则使用
                        if (entity.RowStyleSet.CellBackgroundColorDic != null && entity.RowStyleSet.CellBackgroundColorDic.Keys.Contains(head.ColumnIndex))
                        {
                            style.FillForegroundColor = entity.RowStyleSet.CellBackgroundColorDic[head.ColumnIndex];
                            style.FillPattern = FillPattern.SolidForeground;
                            cell.CellStyle = style;
                        }
                        //如果对行设置背景色，则使用
                        else if (entity.RowStyleSet.RowBackgroundColor != null)
                        {
                            style.FillForegroundColor = entity.RowStyleSet.RowBackgroundColor.Value;
                            style.FillPattern = FillPattern.SolidForeground;
                            cell.CellStyle = style;
                        }

                        //如果列有设置字体颜色，则使用
                        if (entity.RowStyleSet.CellFontColorDic != null && entity.RowStyleSet.CellFontColorDic.Keys.Contains(head.ColumnIndex))
                        {
                            IFont font = excelGlobalDTO.Workbook.CreateFont();//创建字体样式
                            font.Color = entity.RowStyleSet.CellFontColorDic[head.ColumnIndex];//设置字体颜色          
                            style.SetFont(font);
                            cell.CellStyle = style;
                        }
                        //如果对行设置字体颜色，则使用
                        else if (entity.RowStyleSet.RowFontColor != null)
                        {
                            IFont font = excelGlobalDTO.Workbook.CreateFont();//创建字体样式
                            font.Color = entity.RowStyleSet.RowFontColor.Value;//设置字体颜色                        
                            style.SetFont(font);
                            cell.CellStyle = style;
                        }                        
                    }
                });
            }
        }
    }
}

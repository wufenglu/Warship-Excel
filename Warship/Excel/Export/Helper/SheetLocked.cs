using Warship.Excel.Model;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Export.Helper
{
    /// <summary>
    /// 锁定
    /// </summary>
    public class SheetLocked<TEntity> where TEntity : ExcelRowModel, new()
    {

        /// <summary>
        /// 设置锁定
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void SetSheetLocked(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            //遍历Sheet实体集合
            foreach (var item in excelGlobalDTO.Sheets)
            {
                //为空判断
                if (item.SheetHeadList == null)
                {
                    continue;
                }

                //获取Excel的Sheet，对Excel设置批注
                ISheet sheet = excelGlobalDTO.Workbook.GetSheetAt(item.SheetIndex);

                //锁定列头
                ICellStyle cellStyle = excelGlobalDTO.Workbook.CreateCellStyle();

                #region 注释代码
                foreach (var head in item.SheetHeadList)
                {
                    //行是否锁定，此处设置的是否锁定只针对未输入行上的列
                    //IRow row = sheet.GetRow(item.StartRowIndex.Value);
                    //ICell cell = row.GetCell(head.ColumnIndex);
                    //ICellStyle cellStyle = cell.CellStyle;
                    //cellStyle.IsLocked = head.IsLocked;
                    //cell.CellStyle = cellStyle;

                    //设置整列样式
                    //sheet.SetDefaultColumnStyle(head.ColumnIndex, cellStyle);

                    //锁定
                    //if (head.IsLocked == true)
                    //{
                    //    sheet.ProtectSheet("MD5");
                    //}
                }
                #endregion

                //遍历行，设置锁定
                for (int j = (item.StartRowIndex.Value) + 1; j <= sheet.LastRowNum; j++)
                {
                    //获取行
                    IRow row = sheet.GetRow(j);

                    //不存在则跳出
                    if (row == null)
                    {
                        continue;
                    }

                    //遍历头部，设置列
                    foreach (var head in item.SheetHeadList)
                    {
                        //获取列
                        ICell cell = row.GetCell(head.ColumnIndex);

                        //判断单元格是否为空
                        if (cell == null)
                        {
                            continue;
                        }
                        if (cell.CellStyle == null)
                        {
                            cell.CellStyle = excelGlobalDTO.Workbook.CreateCellStyle();
                        }
                        cell.CellStyle.IsLocked = head.IsLocked;

                    }
                }

                foreach (var head in item.SheetHeadList)
                {
                    //锁定
                    if (head.IsLocked == true)
                    {
                        sheet.ProtectSheet("MD5");
                    }
                }
            }
        }
    }
}

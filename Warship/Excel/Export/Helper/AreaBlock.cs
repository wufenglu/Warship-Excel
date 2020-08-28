using Warship.Excel.Model;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Warship.Excel.Export.Helper
{
    /// <summary>
    /// 区块设置
    /// </summary>
    /// <typeparam name="TEntity"></typeparam>
    public class AreaBlock<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 设置区块
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void SetAreaBlock(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            //循环遍历设置区块
            foreach (var item in excelGlobalDTO.Sheets)
            {
                //单个Sheet设置区块
                ISheet sheet = excelGlobalDTO.Workbook.GetSheetAt(item.SheetIndex);
                if (item.AreaBlock != null)
                {
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(item.AreaBlock.StartRowIndex, item.AreaBlock.EndRowIndex, item.AreaBlock.StartColumnIndex, item.AreaBlock.EndColumnIndex);
                    sheet.AddMergedRegion(cellRangeAddress);

                    //创建行、列
                    IRow row = sheet.CreateRow(item.AreaBlock.StartRowIndex);
                    ICell cell = row.CreateCell(item.AreaBlock.StartColumnIndex);
                    cell.SetCellValue(item.AreaBlock.Content);

                    //设置列样式
                    ICellStyle cellStyle = excelGlobalDTO.Workbook.CreateCellStyle();
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.WrapText = true;
                    cell.CellStyle = cellStyle;

                    //设置高度
                    if (item.AreaBlock.Height != null)
                    {
                        row.Height = item.AreaBlock.Height.Value;
                    }
                }
            }

        }
    }
}

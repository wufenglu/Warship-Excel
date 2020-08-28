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
    /// 删除行
    /// </summary>
    public class DeleteRow<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 删除行
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void DeleteRows(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            foreach (var item in excelGlobalDTO.Sheets)
            {
                ISheet sheet = excelGlobalDTO.Workbook.GetSheetAt(item.SheetIndex);

                //判断是否为空
                if (item.SheetEntityList == null)
                {
                    continue;
                }

                /*
                 反向排序：目的是从下向上移动
                 避免从上乡下导致变更后的行号跟实体的行号对不上
                 */
                foreach (var entity in item.SheetEntityList.OrderByDescending(o => o.RowNumber))
                {
                    IRow row = sheet.GetRow(entity.RowNumber);
                    if (entity.IsDeleteRow == true && row != null)
                    {
                        /*
                         说明：startRow、endRow从1开始，从开始行至结束行整体向上移动一行
                         n：负数代表向上移动，正数代表向下移动
                         */
                        //sheet.ShiftRows(entity.RowNumber, sheet.LastRowNum, -1);//有bug，向上移动行后批注没有了
                        row.ZeroHeight = true;//隐藏行
                    }
                }
            }
        }
    }
}

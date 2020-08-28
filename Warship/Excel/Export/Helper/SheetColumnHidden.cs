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
    /// 列隐藏
    /// </summary>
    public class SheetColumnHidden<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 设置隐藏
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void SetSheetColumnHidden(ExcelGlobalDTO<TEntity> excelGlobalDTO)
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
                foreach (var head in item.SheetHeadList)
                {
                    if (head.IsHiddenColumn == true)
                    {
                        //sheet.SetColumnHidden(head.ColumnIndex, head.IsHiddenColumn);

                        //设置列隐藏，使用批注的时候，如果调整内容，再次导入会报错
                        sheet.SetColumnWidth(head.ColumnIndex, 20);
                    }
                }
            }
        }
    }
}

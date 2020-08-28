using Warship.Attribute.Model;
using Warship.Excel.Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Export.Helper
{
    /// <summary>
    /// 批注
    /// </summary>
    public class Comment<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 设置批注
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void SetComment(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            foreach (var item in excelGlobalDTO.Sheets)
            {
                //值判断
                if (item.SheetEntityList == null)
                {
                    continue;
                }

                //获取Excel的Sheet，对Excel设置批注
                ISheet sheet = excelGlobalDTO.Workbook.GetSheetAt(item.SheetIndex);
                //批注总数
                int commentCount = 0;
                foreach (var entity in item.SheetEntityList)
                {
                    if (entity.ColumnErrorMessage == null)
                    {
                        continue;
                    }

                    //NPOI目前最大批注个数设置的默认值是1024个,如果超出默认值就会报错
                    if (commentCount >= 1000)
                    {
                        break;
                    }

                    //获取行对象
                    IRow row = sheet.GetRow(entity.RowNumber);

                    //基于属性设置批注
                    var groupPropertyName = entity.ColumnErrorMessage.Where(w => w.PropertyName != null).GroupBy(g => g.PropertyName);
                    foreach (var groupItem in groupPropertyName)
                    {
                        //NPOI目前最大批注个数设置的默认值是1024个,如果超出默认值就会报错
                        if (commentCount >= 1000)
                        {
                            break;
                        }
                        //获取单元格，设置批注
                        ExcelHeadDTO headDto = item.SheetHeadList.Where(n => n.PropertyName == groupItem.Key).FirstOrDefault();
                        ICell cell = row.GetCell(headDto.ColumnIndex);
                        if (cell == null)
                        {
                            cell = row.CreateCell(headDto.ColumnIndex);
                        }
                        //设置批注
                        string errorMsg = string.Join(";", groupItem.Select(s => s.ErrorMessage).Distinct().ToArray());

                        SetCellComment(cell, errorMsg, excelGlobalDTO);
                        commentCount++;
                    }

                    //基于列头设置批注，适用于动态列
                    var groupColumnName = entity.ColumnErrorMessage.Where(w => w.PropertyName == null).GroupBy(g => g.ColumnName);
                    foreach (var groupItem in groupColumnName)
                    {
                        //NPOI目前最大批注个数设置的默认值是1024个,如果超出默认值就会报错
                        if (commentCount >= 1000)
                        {
                            break;
                        }
                        //获取单元格，设置批注
                        ExcelHeadDTO headDto = item.SheetHeadList.Where(n => n.HeadName == groupItem.Key).FirstOrDefault();
                        ICell cell = row.GetCell(headDto.ColumnIndex);
                        if (cell == null)
                        {
                            cell = row.CreateCell(headDto.ColumnIndex);
                        }
                        //设置批注
                        string errorMsg = string.Join(";", groupItem.Select(s => s.ErrorMessage).Distinct().ToArray());
                        SetCellComment(cell, errorMsg, excelGlobalDTO);
                        commentCount++;
                    }
                }
            }
        }

        /// <summary>
        /// 设置批注
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="errorMsg"></param>
        /// <param name="excelGlobalDTO"></param>
        public void SetCellComment(ICell cell, string errorMsg, ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            ICreationHelper facktory = cell.Row.Sheet.Workbook.GetCreationHelper();
            if (cell.CellComment == null)
            {
                //创建批注区域
                IDrawing patr = cell.Row.Sheet.CreateDrawingPatriarch();
                var anchor = facktory.CreateClientAnchor();
                //设置批注区间大小
                anchor.Col1 = cell.ColumnIndex;
                anchor.Col2 = cell.ColumnIndex + 2;
                //设置列
                anchor.Row1 = cell.RowIndex;
                anchor.Row2 = cell.RowIndex + 3;
                cell.CellComment = patr.CreateCellComment(anchor);
            }
            if (excelGlobalDTO.ExcelVersionEnum == ExcelVersionEnum.V2003)
            {
                cell.CellComment.String = new HSSFRichTextString(errorMsg);//2003批注方式
            }
            else
            {
                cell.CellComment.String = new XSSFRichTextString(errorMsg);//2007批准方式
            }
            cell.CellComment.Author = "yank";
        }

        /// <summary>
        /// 清空批注
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void ClearComment(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            foreach (var item in excelGlobalDTO.Sheets)
            {
                //获取Excel的Sheet，对Excel设置批注
                ISheet sheet = excelGlobalDTO.Workbook.GetSheetAt(item.SheetIndex);

                //清空批注
                for (int i = (item.StartRowIndex.Value); i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null)
                    {
                        continue;
                    }
                    for (var j = row.FirstCellNum; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell == null)
                        {
                            continue;
                        }
                        cell.RemoveCellComment();
                    }
                }
            }
        }
    }
}

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
    /// 列类型
    /// </summary>
    public class SheetColumnType<TEntity> where TEntity : ExcelRowModel, new()
    {

        /// <summary>
        /// 设置列类型
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void SetSheetColumnType(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            //遍历Sheet实体集合
            foreach (var item in excelGlobalDTO.Sheets)
            {
                ISheet sheet = excelGlobalDTO.Workbook.GetSheetAt(item.SheetIndex);

                //判断是否为空
                if (item.SheetEntityList == null)
                {
                    continue;
                }

                //遍历行
                foreach (var entity in item.SheetEntityList)
                {
                    IRow row = sheet.GetRow(entity.RowNumber);
                    foreach (var head in item.SheetHeadList)
                    {
                        //获取列
                        ICell cell = row.GetCell(head.ColumnIndex);

                        //设置生成下拉框的行和列
                        //var cellRegions = new NPOI.SS.Util.CellRangeAddressList((item.StartRowIndex.Value) + 1, sheet.LastRowNum, head.ColumnIndex, head.ColumnIndex);
                        var cellRegions = new NPOI.SS.Util.CellRangeAddressList(entity.RowNumber, entity.RowNumber, head.ColumnIndex, head.ColumnIndex);

                        IDataValidationHelper dvHelper = sheet.GetDataValidationHelper();
                        IDataValidationConstraint dvConstraint = null;
                        IDataValidation dataValidate = null;
                        switch (head.ColumnType)
                        {
                            //列类型为文本
                            case Attribute.Enum.ColumnTypeEnum.Text:
                                break;
                            //列类型为日期
                            case Attribute.Enum.ColumnTypeEnum.Date:
                                head.Format = string.IsNullOrEmpty(head.Format) ? "yyyy-MM-dd" : head.Format;
                                dvConstraint = dvHelper.CreateDateConstraint(OperatorType.BETWEEN, DateTime.MinValue.ToString(head.Format), DateTime.MaxValue.ToString(head.Format), head.Format);
                                dataValidate = dvHelper.CreateValidation(dvConstraint, cellRegions);
                                dataValidate.CreateErrorBox("输入不合法", "必须为日期");
                                dataValidate.ShowPromptBox = true;
                                break;
                            //列类型为浮点
                            case Attribute.Enum.ColumnTypeEnum.Decimal:
                                dvConstraint = dvHelper.CreateDecimalConstraint(OperatorType.BETWEEN, "0", "9999999999");
                                dataValidate = dvHelper.CreateValidation(dvConstraint, cellRegions);
                                dataValidate.CreateErrorBox("输入不合法", "必须在0~9999999999之间。");
                                dataValidate.ShowPromptBox = true;
                                break;
                            //列类型为选项
                            case Attribute.Enum.ColumnTypeEnum.Option:

                                List<string> options = null;

                                #region 全局列设置
                                if (item.ColumnOptions != null && item.ColumnOptions.Count > 0)
                                {
                                    string key = null;
                                    //如果头部名称存在则取头部名称（以头部名称设置选项）
                                    if (item.ColumnOptions.Keys.Contains(head.HeadName) == true)
                                    {
                                        key = head.HeadName;
                                    }
                                    //如果属性存在则取头部名称（以属性名称设置选项）
                                    if (item.ColumnOptions.Keys.Contains(head.PropertyName) == true)
                                    {
                                        key = head.PropertyName;
                                    }

                                    //不为空说明存在，则设置选项
                                    if (key != null)
                                    {
                                        options = item.ColumnOptions[key];
                                    }
                                }
                                #endregion

                                #region 单个列设置
                                //行的优先级高于Sheet的优先级
                                if (entity.ColumnOptions != null && entity.ColumnOptions.Count > 0)
                                {
                                    string key = null;
                                    //如果头部名称存在则取头部名称（以头部名称设置选项）
                                    if (entity.ColumnOptions.Keys.Contains(head.HeadName) == true)
                                    {
                                        key = head.HeadName;
                                    }
                                    //如果属性存在则取头部名称（以属性名称设置选项）
                                    if (entity.ColumnOptions.Keys.Contains(head.PropertyName) == true)
                                    {
                                        key = head.PropertyName;
                                    }

                                    //不为空说明存在，则设置选项
                                    if (key != null)
                                    {
                                        options = entity.ColumnOptions[key];
                                    }
                                }
                                #endregion

                                //不符合条件则跳出
                                if (options == null)
                                {
                                    continue;
                                }
                                if (options.Count() > 0)
                                {
                                    dvConstraint = dvHelper.CreateExplicitListConstraint(options.ToArray());
                                    dataValidate = dvHelper.CreateValidation(dvConstraint, cellRegions);
                                    dataValidate.CreateErrorBox("输入不合法", "请选择下拉列表中的值。");
                                    dataValidate.ShowPromptBox = true;
                                }
                                break;
                        }

                        //类型在指定的范围内是才设置校验
                        if (dataValidate != null)
                        {
                            sheet.AddValidationData(dataValidate);
                        }
                    }
                }
            }
        }
    }
}

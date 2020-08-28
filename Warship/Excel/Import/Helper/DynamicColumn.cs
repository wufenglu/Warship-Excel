using Warship.Attribute.Enum;
using Warship.Attribute.Validations;
using Warship.Excel.Common;
using Warship.Excel.Model;
using Warship.Excel.Model.Column;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using Warship.Utility;

namespace Warship.Excel.Import.Helper
{
    /// <summary>
    /// 验证模式
    /// </summary>
    public enum ValidationModelEnum {
        /// <summary>
        /// 配置列
        /// </summary>
        ConfigColumn = 0,
        /// <summary>
        /// 动态列
        /// </summary>
        DynamicColumn = 1
    };

    /// <summary>
    /// 动态列
    /// </summary>
    public class DynamicColumn<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 动态列验证
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        /// <param name="validationModelEnum"></param>
        public static void ValidationHead(ExcelGlobalDTO<TEntity> excelGlobalDTO, ValidationModelEnum validationModelEnum = ValidationModelEnum.ConfigColumn)
        {
            //为空判断
            if (excelGlobalDTO == null || excelGlobalDTO.Sheets == null)
            {
                return;
            }
            
            //遍历Sheet实体集合
            foreach (var sheetModel in excelGlobalDTO.Sheets)
            {
                ISheet sheet = excelGlobalDTO.Workbook.GetSheetAt(sheetModel.SheetIndex);
                
                //头部校验
                IRow row = sheet.GetRow(sheetModel.StartRowIndex.Value);
                if (row == null)
                {
                    continue;
                }

                //获取表头信息
                List<string> cellValues = row.Cells.Select(s => ExcelHelper.GetCellValue(s)).ToList();
                if (cellValues == null)
                {
                    continue;
                }

                //获取列的设置信息
                List<ColumnModel> columnList = null;
                if (validationModelEnum == ValidationModelEnum.ConfigColumn)
                {
                    columnList = sheetModel.ColumnConfig;
                }
                else
                {
                    columnList = System.Activator.CreateInstance<TEntity>().GetDynamicColumns();
                }

                //当为空的时候跳出,场景：
                if (columnList == null)
                {
                    continue;
                }

                //遍历配置列
                foreach (ColumnModel columnModel in columnList)
                {
                    //如果无校验则跳过
                    if (columnModel.ColumnValidations == null)
                    {
                        continue;
                    }

                    //遍历校验
                    foreach (var validation in columnModel.ColumnValidations)
                    {
                        //校验必填的，判断表头是否在excel中存在
                        if (validation.ValidationTypeEnum == ValidationTypeEnum.Required && cellValues.Contains(columnModel.ColumnName) == false)
                        {
                            throw new Exception(excelGlobalDTO.ExcelValidationMessage.Clgyl_Common_Import_TempletError);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 动态列验证
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        /// <param name="validationModelEnum"></param>
        public static void ValidationValue(ExcelGlobalDTO<TEntity> excelGlobalDTO, ValidationModelEnum validationModelEnum=ValidationModelEnum.ConfigColumn) {

            //为空判断
            if (excelGlobalDTO == null || excelGlobalDTO.Sheets == null) {
                return;
            }
            
            //遍历Sheet实体集合
            foreach (var sheet in excelGlobalDTO.Sheets)
            {
                //获取Sheet头部实体集合
                var headDtoList = sheet.SheetHeadList;
                foreach (var item in sheet.SheetEntityList)
                {
                    //获取列的设置信息
                    List<ColumnModel> columnList = null;
                    if (validationModelEnum == ValidationModelEnum.ConfigColumn)
                    {
                        columnList = sheet.ColumnConfig;
                    }
                    else {
                        columnList = System.Activator.CreateInstance<TEntity>().GetDynamicColumns();
                    }

                    //当为空的时候跳出,场景：
                    if (columnList == null)
                    {
                        continue;
                    }

                    foreach (var config in columnList)
                    {
                        //Excel获取的动态列
                        ColumnModel columnModel = item.OtherColumns.Where(n => n.ColumnName == config.ColumnName).FirstOrDefault();
                        #region 类型校验
                        try
                        {
                            switch (config.ColumnType)
                            {
                                case ColumnTypeEnum.Decimal:
                                    Convert.ToDecimal(columnModel.ColumnValue);
                                    break;
                                case ColumnTypeEnum.Date:
                                    Convert.ToDateTime(columnModel.ColumnValue);
                                    break;
                                case ColumnTypeEnum.DateTime:
                                    Convert.ToDateTime(columnModel.ColumnValue);
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            //异常信息
                            ColumnErrorMessage errorMsg = new ColumnErrorMessage()
                            {
                                ColumnName = config.ColumnName
                            };
                            if (config.ColumnConvertError != null)
                            {
                                errorMsg.ErrorMessage = config.ColumnConvertError.ErrorMessage;
                            }
                            else
                            {
                                errorMsg.ErrorMessage = ex.Message;
                            }
                            //添加至集合
                            item.ColumnErrorMessage.Add(errorMsg);
                        }
                        #endregion

                        if (config.ColumnValidations == null)
                        {
                            continue;
                        }

                        //遍历列校验
                        foreach (ColumnValidationModel validation in config.ColumnValidations)
                        {                            
                            //是否有错误
                            bool isError = true;
                            string errorMessage = null;
                            switch (validation.ValidationTypeEnum)
                            {
                                //必填校验
                                case ValidationTypeEnum.Required:
                                    isError = string.IsNullOrEmpty(columnModel.ColumnValue) ? true : false;
                                    errorMessage = validation.RequiredAttribute.ErrorMessage;
                                    break;
                                //长度校验
                                case ValidationTypeEnum.Length:
                                    isError = columnModel.ColumnValue?.Length > validation.LengthAttribute?.Length ? true : false;
                                    errorMessage = validation.LengthAttribute.ErrorMessage;
                                    break;
                                //范围校验
                                case ValidationTypeEnum.Range:
                                    errorMessage = validation.RangeAttribute.ErrorMessage;
                                    RangeAttributeValidation<TEntity> rangeAttributeValidation = new RangeAttributeValidation<TEntity>();
                                    isError = rangeAttributeValidation.CheckIsError(validation.RangeAttribute, columnModel.ColumnValue);
                                    break;
                                //格式校验
                                case ValidationTypeEnum.Format:
                                    #region 格式校验

                                    errorMessage = validation.FormatAttribute.ErrorMessage;
                                    FormatAttributeValidation<TEntity> formatAttributeValidation = new FormatAttributeValidation<TEntity>();
                                    isError = formatAttributeValidation.CheckIsError(validation.FormatAttribute, columnModel.ColumnValue);

                                    #endregion
                                    break;
                            }

                            //是否有异常
                            if (isError == true)
                            {
                                //异常信息
                                ColumnErrorMessage errorMsg = new ColumnErrorMessage()
                                {
                                    ColumnName = columnModel.ColumnName,
                                    ErrorMessage = errorMessage
                                };
                                //添加至集合
                                item.ColumnErrorMessage.Add(errorMsg);
                            }

                        }
                    }
                }
            }
        }
    }
}

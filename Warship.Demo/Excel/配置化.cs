using Warship.Attribute.Attributes;
using Warship.Attribute.Enum;
using Warship.Excel.Export;
using Warship.Excel.Import;
using Warship.Excel.Model;
using Warship.Excel.Model.Column;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Warship.Demo
{
    class 配置化
    {
        /// <summary>
        /// 动态列
        /// </summary>
        public void ExportByDynamicColumn()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\合同材料导入模版-配置化.xls";

            //代码注释
            //代码注释
            ColumnModel columnModel = new ColumnModel();
            columnModel.ColumnName = "材料名称";
            columnModel.ColumnValidations = new List<ColumnValidationModel>();
            columnModel.ColumnValidations.Add(new ColumnValidationModel()
            {
                ValidationTypeEnum = ValidationTypeEnum.Required,
                RequiredAttribute = new RequiredAttribute() { ErrorMessage = "材料名称必填" }
            });
            columnModel.ColumnValidations.Add(new ColumnValidationModel()
            {
                ValidationTypeEnum = ValidationTypeEnum.Length,
                LengthAttribute = new LengthAttribute(256) { ErrorMessage = "长度必须小于或等于256" }
            });

            //代码注释
            //代码注释
            ColumnModel columnCountModel = new ColumnModel();
            columnCountModel.ColumnName = "数量";
            columnCountModel.ColumnValidations = new List<ColumnValidationModel>();
            columnCountModel.ColumnValidations.Add(new ColumnValidationModel()
            {
                ValidationTypeEnum = ValidationTypeEnum.Required,
                RequiredAttribute = new RequiredAttribute() { ErrorMessage = "数量必填" }
            });

            //代码注释
            //代码注释
            ImportByConfig<ExcelRowModel> import = new ImportByConfig<ExcelRowModel>(1);
            import.ExcelGlobalDTO.SetDefaultSheet("合同材料");


            ExcelSheetModel<ExcelRowModel> sheetModel = import.ExcelGlobalDTO.Sheets.FirstOrDefault();

            //代码注释
            //代码注释
            sheetModel.ColumnConfig = new List<ColumnModel>() {
                columnModel,
                columnCountModel
            };
            import.ExcelGlobalDTO.Sheets.Add(sheetModel);

            //代码注释
            //代码注释
            import.Execute(excelPath);

            //代码注释
            Export<ExcelRowModel> export = new Export<ExcelRowModel>();
            export.Execute(import.ExcelGlobalDTO);
        }
    }
}

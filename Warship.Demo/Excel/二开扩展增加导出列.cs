using Warship.Attribute.Attributes;
using Warship.Attribute.Enum;
using Warship.Excel.Export;
using Warship.Excel.Import;
using Warship.Excel.Model.Column;
using System.Collections.Generic;
using System.IO;

namespace Warship.Demo.Excel
{
    class 二开扩展增加导出列
    {
        /// <summary>
        /// 执行
        /// </summary>
        public void Execute()
        {
            string excelDynamicPath = Directory.GetCurrentDirectory() + "\\..\\Template\\动态列.xls";
            new DynamicColumnDTO().SetDynamicColumns(new List<ColumnModel> {
                new ColumnModel{
                    ColumnName = "动态列A",
                    ColumnValidations=new List<ColumnValidationModel>{
                        new ColumnValidationModel()
                        {
                            ValidationTypeEnum = ValidationTypeEnum.Required,
                            RequiredAttribute = new RequiredAttribute() { ErrorMessage = "动态列A必填" }
                        }
                    }
                },
                new ColumnModel{
                    ColumnName = "动态列B",
                    ColumnValidations=new List<ColumnValidationModel>{
                        new ColumnValidationModel()
                        {
                            ValidationTypeEnum = ValidationTypeEnum.Required,
                            RequiredAttribute = new RequiredAttribute() { ErrorMessage = "动态列B必填" }
                        }
                    }
                }
            });

            Import<DynamicColumnDTO> import = new Import<DynamicColumnDTO>();
            import.ExcelGlobalDTO.SetDefaultSheet("合同材料");
            import.Execute(excelDynamicPath);

            Export<DynamicColumnDTO> export = new Export<DynamicColumnDTO>();
            export.Execute(import.ExcelGlobalDTO);

            //var valuexx = typeof(ContractImportDTO).BaseType.GetProperty("DynamicColumns").GetValue(null, null);

            //ExcelGlobalDTO<ContractImportDTO> globalDTO = new ExcelGlobalDTO<ContractImportDTO>();
            //globalDTO.SetDefaultSheet();
            //globalDTO.FilePath = excelDynamicPath;


        }
    }
}

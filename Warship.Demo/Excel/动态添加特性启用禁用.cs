using Warship.Attribute;
using Warship.Attribute.Attributes;
using Warship.Attribute.Enum;
using Warship.Excel.Export;
using Warship.Excel.Import;
using System.Collections.Generic;
using System.IO;

namespace Warship.Demo.Excel
{
    class 动态添加特性启用禁用
    {
        /// <summary>
        /// 执行
        /// </summary>
        public void Execute()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\动态添加特性启用禁用.xls";

            //设置属性特性
            Dictionary<string, BaseAttribute> dic = new Dictionary<string, BaseAttribute>
            {
                { "Name", new RequiredAttribute { ErrorMessage = "名称必填***" } },
                { "Code", new RequiredAttribute { ErrorMessage = "编码必填***" } }
            };

            //TODO执行启用特性:待优化（统一设置特性，不用单独一个个特性设置）
            AttributeFactory<DemoDTO.ContractImportDTO>.GetValication(ValidationTypeEnum.Required).EnableAttributes(dic);

            //执行禁用特性
            AttributeFactory<DemoDTO.ContractImportDTO>.GetValication(ValidationTypeEnum.Required).DisableAttributes(new List<string> {
                "Name","Code"
            });

            //导入
            Import<DemoDTO.ContractImportDTO> import = new Import<DemoDTO.ContractImportDTO>(1);
            import.Execute(excelPath);

            //导出
            Export<DemoDTO.ContractImportDTO> export = new Export<DemoDTO.ContractImportDTO>();
            export.Execute(import.ExcelGlobalDTO);
        }
    }
}

using Warship.Attribute.Enum;
using Warship.Excel;
using Warship.Excel.Export;
using Warship.Excel.Import;
using Warship.Excel.Model;
using Warship.Excel.Model.Column;
using Warship.Demo;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Demo
{
    public class 标准导入导出
    {
        public void Execute()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\动态添加特性启用禁用.xlsx";

            //导入
            Import<Warship.Demo.DemoDTO.ContractImportDTO> import = new Import<Warship.Demo.DemoDTO.ContractImportDTO>(1);
            import.Execute(excelPath);

            import.ExcelGlobalDTO.FileName = "动态添加特性启用禁用.xlsx";
            Export<Warship.Demo.DemoDTO.ContractImportDTO> export = new Export<Warship.Demo.DemoDTO.ContractImportDTO>();
            export.ExportMemoryStream(import.ExcelGlobalDTO);
        }
    }
}

using Warship.Attribute.Attributes;
using Warship.Attribute.Enum;
using Warship.Excel.Export;
using Warship.Excel.Import;
using Warship.Excel.Model.Column;
using System.Collections.Generic;
using System.IO;
using Warship.Excel.Model;
using Warship.Utility;
using Warship.Demo.DemoDTO;

namespace Warship.Demo.Excel
{
    /// <summary>
    /// 导入注入
    /// </summary>
    public class ImportInjection : IImport<DemoDTO.ContractImportDTO>
    {
        public void ValidationHeaderAfter(ExcelGlobalDTO<DemoDTO.ContractImportDTO> ExcelGlobalDTO)
        {
            return;
        }

        public void ValidationValueAfter(ExcelGlobalDTO<DemoDTO.ContractImportDTO> ExcelGlobalDTO)
        {
            return;
        }
    }

    /// <summary>
    /// 导出注入
    /// </summary>
    public class ExportInjection : IExport<DemoDTO.ContractImportDTO>
    {
        public void ExcelHandleAfter(ExcelGlobalDTO<DemoDTO.ContractImportDTO> excelGlobalDTO)
        {
            return;
        }

        public void ExcelHandleBefore(ExcelGlobalDTO<DemoDTO.ContractImportDTO> excelGlobalDTO)
        {
            return;
        }
    }

    class 二开扩展注入
    {
        /// <summary>
        /// 执行
        /// </summary>
        public void Execute()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\合同导入性能测试.xls";

            //导入注入
            ServiceContainer.Register<IImport<DemoDTO.ContractImportDTO>, ImportInjection>();            

            //执行导入
            Import<DemoDTO.ContractImportDTO> import = new Import<DemoDTO.ContractImportDTO>(0);
            import.ExcelGlobalDTO.SetDefaultSheet();
            import.Execute(excelPath);

            //导出注入
            ServiceContainer.Register<IExport<DemoDTO.ContractImportDTO>, ExportInjection>();

            //执行导出
            Export<DemoDTO.ContractImportDTO> export = new Export<DemoDTO.ContractImportDTO>();
            export.Execute(import.ExcelGlobalDTO);
        }
    }
}

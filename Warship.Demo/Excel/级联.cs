using Warship.Excel.Export;
using Warship.Excel.Import;
using System.IO;

namespace Warship.Demo.Excel
{
    class 级联
    {
        /// <summary>
        /// 执行
        /// </summary>
        public void Execute()
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            string dir = System.Environment.CurrentDirectory;

            //级联处理
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\合同导入模板-级联.xls";
            Import<ContractImportDTO> import = new Import<ContractImportDTO>(1);
            import.ExcelGlobalDTO.SetDefaultSheet();
            import.Execute(excelPath);

            Export<ContractImportDTO> export = new Export<ContractImportDTO>();
            export.Execute(import.ExcelGlobalDTO);

            var errors = import.ExcelGlobalDTO.GetColumnErrorMessages();

            return;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Excel.Export;
using Warship.Excel.Import;
using Warship.Excel.Model;

namespace Warship.Demo.Excel
{
    class 行单元格颜色设置
    {
        public void Execute()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\合同导入性能测试.xlsx";
            Import<DemoDTO.ContractImportDTO> import = new Import<DemoDTO.ContractImportDTO>();
            import.ExcelGlobalDTO.SetDefaultSheet();
            import.Execute(excelPath);

            foreach (var item in import.ExcelGlobalDTO.Sheet.SheetEntityList)
            {
                if (item.RowNumber % 10 == 0)
                {
                    item.RowStyleSet = new RowStyleSet()
                    {
                        //RowBackgroundColor = 53,
                        //RowFontColor = 18,
                        CellBackgroundColorDic = new Dictionary<int, short>() { { 3, 20 } },
                        CellFontColorDic = new Dictionary<int, short> { { 3, 11 } }
                    };
                }
            }

            Export<DemoDTO.ContractImportDTO> export = new Export<DemoDTO.ContractImportDTO>();
            export.Execute(import.ExcelGlobalDTO);
        }
    }
}

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
    public class 材料图片导入
    {
        public void Execute()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\图片导入.xlsx";

            //导入
            Import<Warship.Demo.DemoDTO.ProductDTO> import = new Import<Warship.Demo.DemoDTO.ProductDTO>();
            import.ExcelGlobalDTO.DisableSheetIndexs = new List<int> { 1,2,3,4,5 };
            import.Execute(excelPath);

            Export<Warship.Demo.DemoDTO.ProductDTO> export = new Export<Warship.Demo.DemoDTO.ProductDTO>();
            export.Execute(import.ExcelGlobalDTO);
        }
    }
}

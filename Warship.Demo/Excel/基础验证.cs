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
    public class 基础验证
    {
        public void Execute()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\合同导入模板-基础校验.xls";

            Excel<ContractProductImportDTO> excel = new Excel<ContractProductImportDTO>();
            excel.Import.ExcelGlobalDTO.SetDefaultSheet("合同材料");
            excel.Import.Execute(excelPath);
            excel.Export.Execute(excel.Import.ExcelGlobalDTO);
        }
    }
}

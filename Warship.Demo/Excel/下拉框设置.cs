using Warship.Excel.Export;
using Warship.Excel.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Warship.Demo
{
    class 下拉框设置
    {
        public void OptionsSet()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\Excel\\合同材料导入模版.xls";

            //基于数据的导出
            ExcelGlobalDTO<ContractProductImportDTO> excelGlobalDTO = new ExcelGlobalDTO<ContractProductImportDTO>();
            excelGlobalDTO.SetDefaultSheet();
            excelGlobalDTO.GlobalStartRowIndex = 1;

            Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            dic.Add("类型", new List<string>() {
                "工程类A","采购类B"
            });
            excelGlobalDTO.Sheets.First().ColumnOptions = dic;

            //设置导出错误信息
            Export<ContractProductImportDTO> export = new Export<ContractProductImportDTO>();
            export.ExecuteByData(excelGlobalDTO);
        }
    }
}

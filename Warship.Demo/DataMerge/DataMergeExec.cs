using Warship.Excel.Export;
using Warship.Excel.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Warship.Excel.Import;
using Warship.Excel;
using Warship.Demo.DemoDTO;

namespace Warship.Demo.Excel
{
    class DataMergeExec
    {
        /// <summary>
        /// 执行合并数据
        /// </summary>
        public void Execute() {
            string templatePath = Directory.GetCurrentDirectory() + "\\..\\Template\\merge\\工作流ERP-20200507.xlsx";
            string templatePath2 = Directory.GetCurrentDirectory() + "\\..\\Template\\merge\\工作流ERP.xlsx";

            Excel<DataMerge> dataMerge1 = new Excel<DataMerge>();
            dataMerge1.Import = new Import<DataMerge>(1);
            dataMerge1.Import.ExcelGlobalDTO.DisableSheetIndexs = new List<int>{ 1,2, 3, 4, 5 };
            dataMerge1.Import.Execute(templatePath);

            Excel<DataMerge2> dataMerge2 = new Excel<DataMerge2>();
            dataMerge2.Import = new Import<DataMerge2>();
            dataMerge2.Import.ExcelGlobalDTO.DisableSheetIndexs = new List<int> { 0, 2 ,3,4,5};
            dataMerge2.Import.Execute(templatePath);

            foreach (var item in dataMerge1.Import.ExcelGlobalDTO.Sheet.SheetEntityList) {
                var entity = dataMerge2.Import.ExcelGlobalDTO.Sheet.SheetEntityList.FirstOrDefault(f => f.ID == item.ID);
                if (entity != null) {
                    item.Area = entity.Area;
                    item.Admin = entity.Admin;
                    item.AdminEmail = entity.AdminEmail;
                }
            }
            //dataMerge1.Import.ExcelGlobalDTO.FilePath = templatePath2;
            //dataMerge1.Import.ExcelGlobalDTO.Workbook = null;
            //dataMerge1.Import.ExcelGlobalDTO.Sheet.SheetEntityList = dataMerge1.Import.ExcelGlobalDTO.Sheet.SheetEntityList.OrderByDescending(o => o.WfVersion).ToList();
            dataMerge1.Export.ExecuteByData(dataMerge1.Import.ExcelGlobalDTO);
        }
    }
}

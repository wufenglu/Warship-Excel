using Warship.Excel.Export;
using Warship.Excel.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Warship.Demo
{
    class 基于数据导出
    {
        /// <summary>
        /// 基于数据导出
        /// </summary>
        public void ExportByData()
        {
            List<ContractProductImportDTO> data = new List<ContractProductImportDTO>();
            //代码注释
            for (int i = 0; i < 10; i++)
            {

                //代码注释
                ContractProductImportDTO entity = new ContractProductImportDTO();
                entity.Name = "Name";
                entity.Price = 788;
                entity.Count = 89;

                //代码注释
                data.Add(entity);
            }

            #region 基于数据的导出-单Sheet

            ////准备
            //ExcelGlobalDTO<ContractProductImportDTO> excelGlobalDTO = new ExcelGlobalDTO<ContractProductImportDTO>();
            //excelGlobalDTO.SetDefaultSheet();
            //excelGlobalDTO.Sheets.First().SheetEntityList = data;
            //excelGlobalDTO.FilePath = Environment.CurrentDirectory + "/Template/合同导入模板-导出模板.xls";

            ////导出
            //Export<ContractProductImportDTO> export = new Export<ContractProductImportDTO>();
            //export.ExecuteByData(excelGlobalDTO);

            #endregion

            #region 基于数据的导出-多Sheet

            //基于数据的导出
            ExcelGlobalDTO<ContractProductImportDTO> excelGlobalDTO = new ExcelGlobalDTO<ContractProductImportDTO>();
            excelGlobalDTO.SetDefaultSheet();
            excelGlobalDTO.Sheets.First().SheetEntityList = data;

            //导出第一个
            Export<ContractProductImportDTO> export = new Export<ContractProductImportDTO>();
            export.ExecuteByData(excelGlobalDTO);

            //代码注释
            ExcelGlobalDTO<ContractProductImportDTO> excelGlobalDTO2 = new ExcelGlobalDTO<ContractProductImportDTO>();
            excelGlobalDTO2.Workbook = excelGlobalDTO.Workbook;
            excelGlobalDTO2.SetDefaultSheet();
            excelGlobalDTO2.Sheets.First().SheetEntityList = data;
            excelGlobalDTO2.FilePath = Environment.CurrentDirectory + "/Template/合同导入模板-导出模板.xls";

            //导出第二个
            Export<ContractProductImportDTO> export2 = new Export<ContractProductImportDTO>();
            export2.ExecuteByData(excelGlobalDTO2);

            #endregion
        }
    }
}

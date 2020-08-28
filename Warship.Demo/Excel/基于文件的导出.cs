using Warship.Excel.Export;
using Warship.Excel.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Warship.Demo.Excel
{
    class 基于文件的导出
    {
        /// <summary>
        /// 执行
        /// </summary>
        public void Execute() {
            //级联处理
            string exportTemplatePath = Directory.GetCurrentDirectory() + "\\..\\Template\\合同导入模板-导出模板.xls";

            //第一步
            ExcelGlobalDTO<ContractImportDTO> global = new ExcelGlobalDTO<ContractImportDTO>(1);
            global.SetDefaultSheet("合同");
            global.FilePath = exportTemplatePath;

            //构建数据
            ContractImportDTO model = new ContractImportDTO()
            {
                Name = "A",
                Code = "B"
            };
            global.Sheets.First().SheetEntityList = new List<ContractImportDTO>() {
                model
            };

            //第一步：导出（渲染Excel）
            Export<ContractImportDTO> exprotFile = new Export<ContractImportDTO>();
            exprotFile.ExecuteByFile(exportTemplatePath, ExcelVersionEnum.V2007, global);

            //第二部

            //exprotFile.ExecuteByData(global);

            //第三步：导出（构建文件流）
            exprotFile.ExportMemoryStream(global);
        }
    }
}

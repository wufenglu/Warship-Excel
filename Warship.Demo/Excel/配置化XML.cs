using Warship.Attribute.Attributes;
using Warship.Attribute.Enum;
using Warship.Excel.Export;
using Warship.Excel.Import;
using Warship.Excel.Model;
using Warship.Excel.Model.Column;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using Warship.Utility;
using System;

namespace Warship.Demo
{
    class 配置化XML
    {
        /// <summary>
        /// 动态列
        /// </summary>
        public void Execute()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\合同材料导入模版-配置化XML.xlsx";
            string path = Directory.GetCurrentDirectory() + "\\..\\Template\\合同基本信息.xml";

            //代码注释
            //代码注释
            ImportByConfig<ExcelRowModel> import = new ImportByConfig<ExcelRowModel>(path, 1);
            import.Execute(excelPath);

            EntityProxy.GenerateProxy(import.ExcelGlobalDTO);

            //代码注释
            Export<ExcelRowModel> export = new Export<ExcelRowModel>();
            export.Execute(import.ExcelGlobalDTO);
        }
    }
}

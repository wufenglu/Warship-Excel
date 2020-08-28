using Warship.Excel.Export;
using Warship.Excel.Model;
using System.IO;
using System.Text;

namespace Warship.Demo
{
    class 添加头部说明
    {
        public void AddArea() {
            string excelPath = Directory.GetCurrentDirectory() + "\\Excel\\合同材料导入模版.xls";

            //基于数据的导出
            ExcelGlobalDTO<ContractProductImportDTO> excelGlobalDTO = new ExcelGlobalDTO<ContractProductImportDTO>();
            excelGlobalDTO.SetDefaultSheet();
            excelGlobalDTO.GlobalStartRowIndex = 1;

            //循环设置区块内容
            foreach (var item in excelGlobalDTO.Sheets)
            {
                //设置区块
                item.AreaBlock = new AreaBlock();
                item.AreaBlock.StartRowIndex = 0;
                item.AreaBlock.EndRowIndex = 0;

                //设置区块
                item.AreaBlock.StartColumnIndex = 0;
                item.AreaBlock.EndColumnIndex = 6;
                item.AreaBlock.Height = 256 * 3;

                //设置区块内容
                StringBuilder noteString = new StringBuilder("相关数据字典：（★★请严格按照相关格式填写，以免导入错误★★）\n");
                noteString.Append("1.列名带有' * '是必填列;\n");
                noteString.Append("2.会员卡号：会员卡号长度为3~20位,且只能数字或者英文字母;\n");

                //设置区块内容
                noteString.Append("3.性别：填写“男”或者“女”;\n");

                //设置区块内容
                noteString.Append("4.手机号码：只能是11位数字的标准手机号码;\n");
                noteString.Append("5.固定电话：最好填写为“区号+电话号码”，例：075529755361;\n");

                //设置区块内容
                noteString.Append("6.会员生日：填写格式“年-月-日”，例：1990-12-27，没有则不填;\n");

                //设置区块
                item.AreaBlock.Content = noteString.ToString();
            }

            //设置导出错误信息
            Export<ContractProductImportDTO> export = new Export<ContractProductImportDTO>();
            export.ExecuteByData(excelGlobalDTO);
        }
    }
}

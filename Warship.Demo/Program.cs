using Warship.Attribute;
using Warship.Attribute.Attributes;
using Warship.Attribute.Enum;
using Warship.Attribute.Model;
using Warship.Excel;
using Warship.Excel.Common;
using Warship.Excel.Export;
using Warship.Excel.Import;
using Warship.Excel.Model;
using Warship.Excel.Model.Column;
using Warship.Demo;
using Warship.Demo.Excel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Warship.Utility;
using System.Linq.Expressions;
using Mysoft.Clgyl.Demo;
using System.Threading;

namespace Warship.Demo
{
    class Program
    {
        public static Dictionary<int, List<PerformanceDtlEntity>> dicList = new Dictionary<int, List<PerformanceDtlEntity>>();
        /// <summary>
        /// 基于文件的导出
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {

            Thread thread=new Thread(() => {
                LinkLevel1 level1 = new LinkLevel1();
                level1.Exec();

                var xxx = PerformanceMonitoringLink.GetLinkPerformanceMonitoring();
                int id = Thread.CurrentThread.ManagedThreadId;
                Program.dicList.Add(id, xxx);
            });
            thread.Start();

            Thread thread2 = new Thread(() => {
                LinkLevel1 level1 = new LinkLevel1();
                level1.Exec();

                var xxx = PerformanceMonitoringLink.GetLinkPerformanceMonitoring();
                int id = Thread.CurrentThread.ManagedThreadId;
                Program.dicList.Add(id, xxx);
            });
            thread2.Start();

            Thread thread3 = new Thread(() => {
                LinkLevel1 level1 = new LinkLevel1();
                level1.Exec();

                var xxx = PerformanceMonitoringLink.GetLinkPerformanceMonitoring();
                int id = Thread.CurrentThread.ManagedThreadId;
                Program.dicList.Add(id, xxx);
            });
            thread3.Start();

            Thread thread7 = new Thread(() => {
                LinkLevel1 level1 = new LinkLevel1();
                level1.Exec();

                var xxx = PerformanceMonitoringLink.GetLinkPerformanceMonitoring();
                int id = Thread.CurrentThread.ManagedThreadId;
                Program.dicList.Add(id, xxx);
            });
            thread7.Start();

            Thread thread4 = new Thread(() => {
                LinkLevel1 level1 = new LinkLevel1();
                level1.Exec();

                var xxx = PerformanceMonitoringLink.GetLinkPerformanceMonitoring();
                int id = Thread.CurrentThread.ManagedThreadId;
                Program.dicList.Add(id, xxx);
            });
            thread4.Start();

            Thread thread5 = new Thread(() => {
                LinkLevel1 level1 = new LinkLevel1();
                level1.Exec();

                var xxx = PerformanceMonitoringLink.GetLinkPerformanceMonitoring();
                int id = Thread.CurrentThread.ManagedThreadId;
                Program.dicList.Add(id, xxx);
            });
            thread5.Start();

            Thread thread6 = new Thread(() => {
                LinkLevel1 level1 = new LinkLevel1();
                level1.Exec();

                var xxx = PerformanceMonitoringLink.GetLinkPerformanceMonitoring();
                int id = Thread.CurrentThread.ManagedThreadId;
                Program.dicList.Add(id, xxx);
            });
            thread6.Start();

            Thread.Sleep(10000);
            return;


            //标准导入导出 model = new 标准导入导出();
            //model.Execute();

            //ContractImportDTO dto = new ContractImportDTO();
            //ValidationHelper.Exec(dto);

            //DataMergeExec dataMergeExec = new DataMergeExec();
            //dataMergeExec.Execute();


            //材料图片导入 model = new 材料图片导入();
            //model.Execute();

            //行单元格颜色设置 model = new 行单元格颜色设置();
            //model.Execute();

            //配置化XML model = new 配置化XML();
            //model.Execute();

            //二开扩展注入 model = new 二开扩展注入();
            //model.Execute();

            //基于数据导出 model = new 基于数据导出();
            //model.ExportByData();

            //二开扩展增加导出列 model2 = new 二开扩展增加导出列();
            //model.Execute();

            //级联 model = new 级联();
            //model.Execute();

            //二开扩展增加导出列 model = new 二开扩展增加导出列();
            //model.Execute();

            //配置化 dynamic = new 配置化();
            //dynamic.ExportByDynamicColumn();

            //动态添加特性启用禁用 attribute = new 动态添加特性启用禁用();
            //attribute.Execute();

            //获取所有列的错误信息 attribute = new 获取所有列的错误信息();
            //attribute.Execute();

        }


        //string url = System.Web.HttpUtility.UrlDecode("http%3A%2F%2F10.5.11.101%3A10120%2F%2FMyWorkflow%2FWF_ProcessInitiate_Form_Transfer.aspx%3Fmode%3D1%26opentype%3Dbizsys%26processguid%3D%26businessGUID%3Db333c498-71d6-e811-80c3-00155d0a444e%26businessType%3D%25u6750%25u6599%25u7533%25u8BF7%25u5BA1%25u6279");

        //string encodeUrl = System.Web.HttpUtility.UrlEncode("http://10.5.11.101:10120//MyWorkflow/WF_ProcessInitiate_Form_Transfer.aspx?mode=1&opentype=bizsys&processguid=&businessGUID=b333c498-71d6-e811-80c3-00155d0a444e&businessType=材料申请审批");


        /// <summary>
        /// 级联
        /// </summary>
        /// <param name="args"></param>
        void Test(string[] args)
        {
            ContractImportDTO contract = new ContractImportDTO();
            PropertyInfo prop = contract.GetType().GetProperty("Products");
            bool isGenericType = prop.PropertyType.IsGenericType;
            Type[] types = prop.PropertyType.GetGenericArguments();
            Type type = types[0];

            string assemblyFile = "Warship.dll";
            string className = "Warship.Excel.Model.ExcelSheetModel`1";// + "<" + type.FullName + ">";
            string productAssemblyName = type.Assembly.GetName().Name;
            //string className = "Warship.Excel.Model.ExcelRowModel";
            //string className = type.FullName;
            //Assembly.Load(type.FullName);

            //Type testType = Assembly.Load("Warship").GetType(className).MakeGenericType(type); ;

            //加载泛型类
            Type genericClass = Assembly.LoadFrom(assemblyFile).GetType(className);
            //制作泛型类型
            genericClass = genericClass.MakeGenericType(type);

            //ExcelSheetModel < typeof(ExcelRowModel) > propSheetModel = new ExcelSheetModel<TEntity>();

            //ExcelSheetModel<ContractProductImportDTO> propSheetModel = System.Activator.CreateInstance(Type.GetType(className)) as ExcelSheetModel<ContractProductImportDTO>;
            var propSheetModel = System.Activator.CreateInstance(genericClass);

            SheetAttribute sheetAttribute = prop.GetCustomAttribute<SheetAttribute>();
            if (sheetAttribute != null)
            {
                ////获取Sheet信息
                //ISheet propSheet = ExcelGlobalDTO.Workbook.GetSheet(sheetAttribute.SheetName);
                //if (prop.PropertyType.IsGenericType)
                //{
                //    Type[] types = prop.PropertyType.GetGenericArguments();
                //    Type type = prop.PropertyType.GetGenericArguments()[0];


                //构建SheetModel
                //ExcelSheetModel <typeof(ExcelRowModel)> propSheetModel = new ExcelSheetModel<TEntity>();
                //propSheetModel.StartRowIndex = sheetAttribute.StartRowIndex;
                //propSheetModel.StartColumnIndex = sheetAttribute.StartColumnIndex;
                //propSheetModel.SheetIndex = ExcelGlobalDTO.Workbook.GetSheetIndex(sheetAttribute.SheetName);
                //propSheetModel.SheetName = sheetAttribute.SheetName;
                //}
                ////TODO
            }

            //变量声明
            DateTime dt = new DateTime();
            DateTime.TryParse("12-9月-2018", out dt);
            DateTime dt1 = Convert.ToDateTime("12-九月-2018");
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\合同导入模板.xls";

            Import<ContractImportDTO> import = new Import<ContractImportDTO>(1);
            import.ExcelGlobalDTO.SetDefaultSheet();
            import.Execute(excelPath);

            Export<ContractImportDTO> export = new Export<ContractImportDTO>();
            export.Execute(import.ExcelGlobalDTO);

            #region 基于实体的导入导出

            ////导入
            //Import<ContractProductImportDTO> import = new Import<ContractProductImportDTO>(1);
            //import.ExcelGlobalDTO.DisableSheetIndexs = new List<int>();
            //import.ExcelGlobalDTO.DisableSheetIndexs.Add(1);
            //import.Execute(excelPath);
            //import.ExcelGlobalDTO.Sheets.First().ColumnOptions = new Dictionary<string, List<string>>();
            //import.ExcelGlobalDTO.Sheets.First().ColumnOptions.Add("类型",new List<string>() {
            //    "工程类","采购类"
            //});

            //Dictionary<string, List<string>> dic = new Dictionary<string, List<string>>();
            //dic.Add("类型", new List<string>() {
            //    "工程类A","采购类B"
            //});
            //import.ExcelGlobalDTO.Sheets.First().SheetEntityList[1].ColumnOptions = dic;

            ////循环设置区块内容
            //foreach (var item in import.ExcelGlobalDTO.Sheets)
            //{
            //    //设置区块
            //    item.AreaBlock = new Component.Office.Excel.Model.AreaBlock();
            //    item.AreaBlock.StartRowIndex = 0;
            //    item.AreaBlock.EndRowIndex = 0;

            //    //设置区块
            //    item.AreaBlock.StartColumnIndex = 0;
            //    item.AreaBlock.EndColumnIndex = 6;
            //    item.AreaBlock.Height = 256 * 3;

            //    //设置区块内容
            //    StringBuilder noteString = new StringBuilder("相关数据字典：（★★请严格按照相关格式填写，以免导入错误★★）\n");
            //    noteString.Append("1.列名带有' * '是必填列;\n");
            //    noteString.Append("2.会员卡号：会员卡号长度为3~20位,且只能数字或者英文字母;\n");

            //    //设置区块内容
            //    noteString.Append("3.性别：填写“男”或者“女”;\n");

            //    //设置区块内容
            //    noteString.Append("4.手机号码：只能是11位数字的标准手机号码;\n");
            //    noteString.Append("5.固定电话：最好填写为“区号+电话号码”，例：075529755361;\n");

            //    //设置区块内容
            //    noteString.Append("6.会员生日：填写格式“年-月-日”，例：1990-12-27，没有则不填;\n");

            //    //设置区块
            //    item.AreaBlock.Content = noteString.ToString();
            //}

            ////设置导出错误信息
            //Export<ContractProductImportDTO> export = new Export<ContractProductImportDTO>();
            //export.Execute(import.ExcelGlobalDTO);            

            #endregion
        }
    }

    public class VEntity
    {
        /// <summary>
        /// 名称
        /// </summary>
        [Required(ErrorMessage = "请输入名称")]
        [Length(20, ErrorMessage = "长度不能超过20个字符")]
        public string Name { get; set; }

        /// <summary>
        /// 邮箱
        /// </summary>
        [Required(ErrorMessage = "请输入邮箱")]
        public string Email { get; set; }
    }
}

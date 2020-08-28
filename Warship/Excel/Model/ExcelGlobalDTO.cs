using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Attribute.Model;
using System.IO;
using Warship.Excel.Model.Const;
using Warship.Attribute.Attributes;
using System.Reflection;
using System.Collections;
using Warship.Utility;

namespace Warship.Excel.Model
{
    /// <summary>
    /// Excel全局对象
    /// </summary>
    public class ExcelGlobalDTO<TEntity> where TEntity : ExcelRowModel
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public ExcelGlobalDTO()
        {
            ExcelValidationMessage = new ExcelValidationMessage();
            PerformanceMonitoring = new PerformanceMonitoring();
            Sheets = new List<ExcelSheetModel<TEntity>>();
        }

        /// <summary>
        /// 初始化起始行、起始列
        /// </summary>
        /// <param name="startRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        public ExcelGlobalDTO(int startRowIndex = 0, int startColumnIndex = 0)
        {
            ExcelValidationMessage = new ExcelValidationMessage();
            PerformanceMonitoring = new PerformanceMonitoring();
            GlobalStartRowIndex = startRowIndex;
            GlobalStartColumnIndex = startColumnIndex;
            Sheets = new List<ExcelSheetModel<TEntity>>();
        }

        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 文件路径：以文件路径处理时使用
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// 文件字节数组
        /// </summary>
        public byte[] FileBytes { get; set; }

        /// <summary>
        /// Excel版本
        /// </summary>
        public ExcelVersionEnum ExcelVersionEnum { get; set; } = ExcelVersionEnum.V2007;

        /// <summary>
        /// 工作簿对象
        /// </summary>
        public IWorkbook Workbook { get; set; }

        /// <summary>
        /// 全局起始行：不设置Sheet的默认按照全局来
        /// </summary>
        public int GlobalStartRowIndex { get; set; }

        /// <summary>
        /// 全局起始列：不设置Sheet的默认按照全局来
        /// </summary>
        public int GlobalStartColumnIndex { get; set; }

        /// <summary>
        /// Sheet集合
        /// </summary>
        public List<ExcelSheetModel<TEntity>> Sheets { get; set; }

        /// <summary>
        /// Sheet，从Sheets中取第一条
        /// </summary>
        public ExcelSheetModel<TEntity> Sheet
        {
            get
            {
                if (Sheets.IsNullOrEmpty())
                {
                    return null;
                }
                else
                {
                    return Sheets.First();
                }
            }
            set
            {
                Sheets = new List<ExcelSheetModel<TEntity>>
                {
                    value
                };
            }
        }

        /// <summary>
        /// 禁用的Sheet的索引，被禁用的sheet页面将不读取
        /// <remarks> 
        /// 在多sheet的execl下，有些sheet页面只是一些辅助信息，是不需要读取的
        /// </remarks>
        /// </summary>
        public List<int> DisableSheetIndexs { get; set; }

        /// <summary>
        /// Excel验证消息
        /// </summary>
        public ExcelValidationMessage ExcelValidationMessage { get; set; }

        /// <summary>
        /// 设置默认的Sheet，适用于基于数据的导出及代码实现配置化（非xml）
        /// </summary>
        /// <param name="sheetName"></param>
        public void SetDefaultSheet(string sheetName = null)
        {
            //Sheet设置
            Sheets = new List<ExcelSheetModel<TEntity>>();
            ExcelSheetModel<TEntity> sheetModel = new ExcelSheetModel<TEntity>
            {
                SheetName = sheetName
            };
            Sheets.Add(sheetModel);
        }
        /// <summary>
        /// 设置活动的execl
        /// </summary>
        /// <remarks>
        /// 将execl中的某一个sheet页设置成为当前活动页，用户打开execl的时候会默认定位到该sheet的页
        ///
        /// </remarks> 
        /// <param name="sheetIndex">sheet页的索引值，从左到右从0开始</param>
        public void SetActiveSheet(int sheetIndex)
        {
            if (Workbook == null)
            {
                return;
            }
            Workbook.SetActiveSheet(sheetIndex);
            Workbook.SetSelectedTab(sheetIndex);
        }

        /// <summary>
        /// 从表的Excel全局对象
        /// </summary>
        public object SlaveExcelGlobalDTO { get; set; }

        /// <summary>
        /// 获取所有行错误信息：{"Sheet1":[],"Sheet2":[]}
        /// </summary>
        public List<ColumnErrorMessage> GetColumnErrorMessages()
        {
            List<ColumnErrorMessage> errors = new List<ColumnErrorMessage>();
            foreach (var item in Sheets)
            {
                if (item.SheetEntityList == null)
                {
                    continue;
                }
                item.SheetEntityList.ForEach(n =>
                {
                    List<ColumnErrorMessage> errorsResult = GetMasterSlaveErrorMessages(n);
                    errors.AddRange(errorsResult);
                });
            }
            return errors;
        }

        /// <summary>
        /// 获取下级的错误信息
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        private List<ColumnErrorMessage> GetMasterSlaveErrorMessages(TEntity entity)
        {
            //获取当前实体的错误信息
            List<ColumnErrorMessage> errors = new List<ColumnErrorMessage>();
            errors.AddRange(entity.ColumnErrorMessage);

            //遍历实体属性，寻找级联实体，找到异常信息
            foreach (var prop in entity.GetType().GetProperties())
            {
                //获取Sheet特性标记
                SheetAttribute sheetAttribute = prop.GetCustomAttribute<SheetAttribute>();
                if (sheetAttribute == null)
                {
                    continue;
                }
                
                //当为级联对象时处理
                var values = prop.GetValue(entity, null) as IEnumerable;
                if (values == null)
                {
                    continue;
                }

                //程序集
                string assemblyString = "Warship";
                string importClassName = "Warship.Excel.Model.ExcelGlobalDTO`1";

                //类型
                Type[] types = prop.PropertyType.GetGenericArguments();
                Type type = prop.PropertyType.GetGenericArguments()[0];

                //值不为空的时遍历实体
                foreach(var value in values)
                {
                    Type importType = Assembly.Load(assemblyString).GetType(importClassName).MakeGenericType(type);
                    object importModel = System.Activator.CreateInstance(importType);
                    BindingFlags flag = BindingFlags.Instance | BindingFlags.NonPublic;
                    List<ColumnErrorMessage> errorsResult = importType.GetMethod("GetMasterSlaveErrorMessages", flag).Invoke(importModel, new object[] { value }) as List<ColumnErrorMessage>;
                    errors.AddRange(errorsResult);
                };
            }
            return errors;
        }

        /// <summary>
        /// 性能对象
        /// </summary>
        public PerformanceMonitoring PerformanceMonitoring { get; set; }
    }
}

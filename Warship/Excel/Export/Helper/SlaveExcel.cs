using Warship.Excel.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Export.Helper
{
    /// <summary>
    /// 主从级联：从Excel处理
    /// </summary>
    public class SlaveExcel<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 主从级联：从Excel处理
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void SlaveExcelHandle(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            //判断是否有从Excel（主从级联）
            if (excelGlobalDTO.SlaveExcelGlobalDTO != null)
            {
                //设置Workbook,不断基于父级写样式
                Type globalType = excelGlobalDTO.SlaveExcelGlobalDTO.GetType();
                globalType.GetProperty("Workbook").SetValue(excelGlobalDTO.SlaveExcelGlobalDTO, excelGlobalDTO.Workbook);
                string className = "Warship.Excel.Export.Export`1";
                Type slaveType = Type.GetType(className).MakeGenericType(globalType.GetGenericArguments()[0]);
                object slaveEntity = Activator.CreateInstance(slaveType);
                slaveType.GetMethod("Execute").Invoke(slaveEntity, new object[] { excelGlobalDTO.SlaveExcelGlobalDTO });
            }
        }
    }
}

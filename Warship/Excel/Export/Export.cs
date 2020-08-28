using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Warship.Attribute;
using Warship.Attribute.Model;
using Warship.Excel.Model;
using System.Web;
using Warship.Excel.Common;
using NPOI.SS.Util;
using Warship.Excel.Export.Helper;
using Warship.Excel.Model.Column;
using System.Reflection;
using Warship.Utility;
using static Warship.Excel.Common.ExcelHelper;


namespace Warship.Excel.Export
{
    /// <summary>
    /// 导出
    /// </summary>
    public class Export<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 表头样式
        /// </summary>
        public Dictionary<string, ICellStyle> HeadCellStyleDic { set; get; }
        /// <summary>
        /// 单元格样式字典，key是表头名称
        /// </summary>
        public Dictionary<string, ICellStyle> CellStyleDic { set; get; }

        #region 导出

        /// <summary>
        /// 基于导入执行导出
        /// </summary>
        public void Execute(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            excelGlobalDTO.PerformanceMonitoring.Start("=========【导出】Excel处理======");

            //Excel处理
            ExcelHandle(excelGlobalDTO);

            //删除行
            DeleteRow<TEntity> deleteRow = new DeleteRow<TEntity>();
            deleteRow.DeleteRows(excelGlobalDTO);

            excelGlobalDTO.PerformanceMonitoring.Stop();

            //如果路径文件为空则不存储
            if (string.IsNullOrEmpty(excelGlobalDTO.FilePath) == true)
            {
                return;
            }

            //写入内容
            using (FileStream fs = new FileStream(excelGlobalDTO.FilePath, FileMode.Create))
            {
                excelGlobalDTO.Workbook.Write(fs);
                fs.Dispose();
                fs.Close();
            }
        }

        /// <summary>
        /// 根据模板导出Excel（适用模板+数据）
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="excelVersionEnum"></param>
        /// <param name="excelGlobalDTO"></param>
        public virtual void ExecuteByFile(string filePath, ExcelVersionEnum excelVersionEnum, ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            //级联处理
            FileInfo fileInfo = new FileInfo(filePath);
            byte[] buffers = new byte[fileInfo.Length];
            using (FileStream fs = new FileStream(filePath, FileMode.Open))
            {
                int read;
                while ((read = fs.Read(buffers, 0, buffers.Length)) > 0)
                {
                    fs.Dispose();
                }
                //fs.Read(buffers, 0, buffers.Length)
            }
            ExecuteByEmptyBuffer(buffers, excelVersionEnum, excelGlobalDTO);
        }

        /// <summary>
        /// 根据模板Excel执行
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="excelVersionEnum"></param>
        /// <param name="excelGlobalDTO"></param>
        public virtual void ExecuteByEmptyBuffer(byte[] buffer, ExcelVersionEnum excelVersionEnum, ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            Stream stream = new MemoryStream(buffer);
            excelGlobalDTO.Workbook = ExcelHelper.GetWorkbook(stream, excelVersionEnum);
            this.ExecuteByData(excelGlobalDTO);
        }

        /// <summary>
        /// 基于数据执行导出：FilePath、SheetEntityList、StartRowIndex
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void ExecuteByData(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            //初始化
            excelGlobalDTO.PerformanceMonitoring.Start("=========【导出】数据初始化======");
            GlobalDataInit(excelGlobalDTO);
            excelGlobalDTO.PerformanceMonitoring.Stop();

            //初始化Excel
            excelGlobalDTO.PerformanceMonitoring.Start("=========【导出】初始化Excel======");
            GlobalExcelInitByData(excelGlobalDTO);
            excelGlobalDTO.PerformanceMonitoring.Stop();

            //Excel处理
            excelGlobalDTO.PerformanceMonitoring.Start("=========【导出】Excel处理======");
            ExcelHandle(excelGlobalDTO);
            excelGlobalDTO.PerformanceMonitoring.Stop();

            //如果路径文件为空则不存储
            if (string.IsNullOrEmpty(excelGlobalDTO.FilePath) == true)
            {
                return;
            }

            //写入内容
            using (FileStream fs = new FileStream(excelGlobalDTO.FilePath, FileMode.Create))
            {
                excelGlobalDTO.Workbook.Write(fs);
                fs.Dispose();
                fs.Close();
            }
        }

        /// <summary>
        /// 执行导出
        /// </summary>
        /// <param name="excelGlobalDTO"> Excel全局对象</param>
        public void ExportMemoryStream(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            //生成名称
            string fileName = excelGlobalDTO.FileName.Split('.')[0]
                + DateTime.Now.ToString("yyyyMMdd")
                + "."
                + excelGlobalDTO.FileName.Split('.')[1];

            fileName = string.Format("attachment;filename='{0}';filename*=utf-8''{0}", HttpUtility.UrlEncode(fileName));

            //当前请求上下文
            HttpResponse response = HttpContext.Current.Response;
            
            response.Clear();
            response.ClearHeaders();
            response.ClearContent();

            response.Buffer = true;
            response.ContentEncoding = System.Text.Encoding.UTF8;
            response.ContentType = "application/vnd.ms-excel";
            response.AddHeader("Content-Disposition", fileName);

            MemoryStream ms = new MemoryStream();
            excelGlobalDTO.Workbook.Write(ms);

            byte[] buffer = ms.ToArray();
            response.AddHeader("Content-Length", buffer.Length.ToString());
            response.BinaryWrite(buffer);

            response.Flush();
            response.Close();
        }

        #endregion

        #region 初始化

        /// <summary>
        /// 基于数据导出的初始化
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        private void GlobalDataInit(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            #region 如果Sheet为空则初始化Sheet
            if (excelGlobalDTO.Sheets == null)
            {
                //构建默认一个
                ExcelSheetModel<TEntity> sheetModel = new ExcelSheetModel<TEntity>
                {
                    SheetIndex = 0,
                    StartRowIndex = excelGlobalDTO.GlobalStartRowIndex,
                    StartColumnIndex = excelGlobalDTO.GlobalStartColumnIndex
                };

                //设置一个默认
                excelGlobalDTO.Sheets = new List<ExcelSheetModel<TEntity>>
                {
                    sheetModel
                };
            }
            else
            {
                //如果未设置起始行起始列，则以全局为准
                foreach (var item in excelGlobalDTO.Sheets)
                {
                    item.StartRowIndex = item.StartRowIndex ?? excelGlobalDTO.GlobalStartRowIndex;
                    item.StartColumnIndex = item.StartColumnIndex ?? excelGlobalDTO.GlobalStartColumnIndex;
                }
            }
            #endregion
        }

        /// <summary>
        /// 初始化Excel
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        private void GlobalExcelInitByData(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            //工作簿为空判断
            if (excelGlobalDTO.Workbook == null)
            {
                excelGlobalDTO.Workbook = ExcelHelper.CreateWorkbook(excelGlobalDTO.ExcelVersionEnum);
            }

            //遍历Sheet，创建头部、数据行
            foreach (var item in excelGlobalDTO.Sheets)
            {
                item.SheetHeadList = ExcelAttributeHelper<TEntity>.GetHeads();
                //获取动态列
                List<ColumnModel> dynamicColumns = System.Activator.CreateInstance<TEntity>().GetDynamicColumns();
                if (dynamicColumns != null && dynamicColumns.Count > 0)
                {
                    if (item.SheetHeadList == null)
                    {
                        item.SheetHeadList = new List<ExcelHeadDTO>();
                    }
                    foreach (ColumnModel columnModel in dynamicColumns)
                    {
                        //头部实体赋值
                        ExcelHeadDTO excelHeadDTO = new ExcelHeadDTO();
                        excelHeadDTO.ColumnIndex = columnModel.ColumnIndex;
                        excelHeadDTO.SortValue = columnModel.SortValue;
                        excelHeadDTO.HeadName = columnModel.ColumnName;
                        excelHeadDTO.IsValidationHead = columnModel.IsValidationHead;
                        excelHeadDTO.IsSetHeadColor = columnModel.IsSetHeadColor;
                        excelHeadDTO.IsLocked = columnModel.IsLocked;
                        excelHeadDTO.ColumnType = columnModel.ColumnType;
                        excelHeadDTO.IsHiddenColumn = columnModel.IsHiddenColumn;
                        excelHeadDTO.ColumnWidth = columnModel.ColumnWidth;
                        excelHeadDTO.Format = columnModel.Format;
                        excelHeadDTO.BackgroundColor = columnModel.BackgroundColor;
                        item.SheetHeadList.Add(excelHeadDTO);
                    }

                    //排序
                    item.SheetHeadList = item.SheetHeadList.OrderBy(o => o.SortValue).ToList();
                    int order = 0;
                    item.SheetHeadList.ForEach(it => { it.ColumnIndex = order++; });
                }

                //创建Sheet
                ISheet sheet = null;
                #region 创建Sheet
                //判断Sheet是否存在，不存在则创建
                if (string.IsNullOrEmpty(item.SheetName))
                {
                    sheet = excelGlobalDTO.Workbook.CreateSheet();
                    item.SheetName = sheet.SheetName;
                }
                else
                {
                    //如果存在则使用存在
                    ISheet existSheet = excelGlobalDTO.Workbook.GetSheet(item.SheetName);
                    if (existSheet == null)
                    {
                        //处理特殊字符
                        string sheetName = item.SheetName.Replace("\\", "").Replace("/", "").Replace(":", "").Replace("?", "")
                            .Replace("*", "").Replace("[", "").Replace("]", "");
                        sheet = excelGlobalDTO.Workbook.CreateSheet(sheetName);
                        item.SheetName = sheet.SheetName;
                    }
                    else
                    {
                        sheet = existSheet;
                    }
                }
                #endregion

                item.SheetIndex = excelGlobalDTO.Workbook.GetSheetIndex(sheet.SheetName);

                //初始化sheet页表头和单元格样式
                InitCellStyle(item.SheetHeadList, excelGlobalDTO.Workbook);

                #region 创建头部

                //创建头部
                bool isExistHead = sheet.GetRow(item.StartRowIndex.Value) == null ? false : true;//是否存在头部
                IRow row = sheet.GetRow(item.StartRowIndex.Value) ?? sheet.CreateRow(item.StartRowIndex.Value);
                Dictionary<string, ICellStyle> styleDic = new Dictionary<string, ICellStyle>();
                //头部
                foreach (var head in item.SheetHeadList)
                {
                    //获取单元格对象
                    ICell cell = row.GetCell(head.ColumnIndex) ?? row.CreateCell(head.ColumnIndex);
                    //设置样式
                    if (cell.CellStyle.IsNullOrEmpty())
                    {
                        cell.CellStyle = HeadCellStyleDic[head.HeadName];
                    }
                    //设置值
                    cell.SetCellValue(head.HeadName);

                    //设置列宽
                    if (head.ColumnWidth == 0 && head.HeadName != null)
                    {
                        head.ColumnWidth = Encoding.Default.GetBytes(head.HeadName).Length;
                    }
                    if (isExistHead == false)
                    {
                        head.ColumnWidth = head.ColumnWidth + 1;//多留一个字符的宽度
                        sheet.SetColumnWidth(head.ColumnIndex, head.ColumnWidth * 256);
                    }
                }
                #endregion

                //如果没有实体集合则跳出
                if (item.SheetEntityList == null)
                {
                    continue;
                }

                //设置行号
                int rowNumber = item.StartRowIndex.Value;
                foreach (var entity in item.SheetEntityList)
                {
                    //设置实体的行号
                    rowNumber++;
                    entity.RowNumber = rowNumber;
                }
                #region 创建行、列及设置值
                foreach (var entity in item.SheetEntityList)
                {
                    //创建行
                    IRow dataRow = sheet.GetRow(entity.RowNumber) ?? sheet.CreateRow(entity.RowNumber);
                    //循环创建列
                    foreach (var head in item.SheetHeadList)
                    {

                        //获取单元格对象
                        ICell cell = dataRow.GetCell(head.ColumnIndex) ?? dataRow.CreateCell(head.ColumnIndex);
                        //获取单元格样式
                        if (cell.CellStyle.IsNullOrEmpty())
                        {
                            cell.CellStyle = CellStyleDic[head.HeadName];
                        }

                        object value = null;
                        if (string.IsNullOrEmpty(head.PropertyName) == false)
                        {
                            value = entity.GetType().GetProperty(head.PropertyName).GetValue(entity);
                        }
                        else
                        {
                            if (entity.OtherColumns != null)
                            {
                                ColumnModel columnModel = entity.OtherColumns.Where(w => w.ColumnName == head.HeadName).FirstOrDefault();
                                value = columnModel?.ColumnValue;
                            }
                        }
                        if (value != null)
                        {
                            if (head.ColumnType == Attribute.Enum.ColumnTypeEnum.Date && string.IsNullOrEmpty(head.Format) == false)
                            {
                                cell.SetCellValue(((DateTime)value).ToString(head.Format));
                            }
                            else if (head.ColumnType == Attribute.Enum.ColumnTypeEnum.Decimal && string.IsNullOrEmpty(head.Format) == false)
                            {
                                string cellValue = ((decimal)value).ToString(GetNewFormatString(head.Format));
                                cell.SetCellValue(cellValue);
                            }
                            else if (head.ColumnType == Attribute.Enum.ColumnTypeEnum.Option && string.IsNullOrEmpty(head.Format) == false)
                            {
                                string cellValue = ((decimal)value).ToString(GetNewFormatString(head.Format));
                                cell.SetCellValue(cellValue);
                            }
                            else
                            {
                                cell.SetCellValue(value.ToString());
                            }
                        }
                    }
                }
                #endregion
            }
        }


        /// <summary>
        /// 初始化表头单元格样式
        /// </summary>
        /// <param name="excelHeadDTOList"></param>
        /// <param name="workbook"></param>
        public void InitCellStyle(List<ExcelHeadDTO> excelHeadDTOList, IWorkbook workbook)
        {
            if (HeadCellStyleDic == null)
            {
                HeadCellStyleDic = new Dictionary<string, ICellStyle>();
            }
            if (CellStyleDic == null)
            {
                CellStyleDic = new Dictionary<string, ICellStyle>();
            }

            foreach (var head in excelHeadDTOList)
            {

                #region 头部样式
                //如果通过属性制定了样式则以用户指定的为主，对象上面的加的样式特性不生效
                if (HeadCellStyleDic.ContainsKey(head.HeadName) == false)
                {
                    ICellStyle headStyle = workbook.CreateCellStyle();
                    headStyle.BorderBottom = BorderStyle.Thin;
                    headStyle.BorderLeft = BorderStyle.Thin;
                    headStyle.BorderRight = BorderStyle.Thin;
                    headStyle.BorderTop = BorderStyle.Thin;
                    headStyle.Alignment = HorizontalAlignment.Left;//列头都是左对齐
                    if (head.HeaderBackgroundColor != 0)
                    {
                        headStyle.FillPattern = FillPattern.SolidForeground;
                        headStyle.FillForegroundColor = head.HeaderBackgroundColor;
                    }
                    HeadCellStyleDic.Add(head.HeadName, headStyle);
                }

                #endregion

                #region 单元格样式
                if (CellStyleDic.ContainsKey(head.HeadName) == false)
                {

                    ICellStyle cellStyle = workbook.CreateCellStyle();
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    cellStyle.Alignment = HorizontalAlignment.Left;
                    if (head.BackgroundColor != 0)
                    {
                        cellStyle.FillPattern = FillPattern.SolidForeground;
                        cellStyle.FillForegroundColor = head.BackgroundColor;
                    }
                    if (string.IsNullOrEmpty(head.Format) == false)
                    {
                        IDataFormat format = workbook.CreateDataFormat();
                        cellStyle.DataFormat = format.GetFormat(GetNewFormatString(head.Format));
                        if (head.ColumnType == Attribute.Enum.ColumnTypeEnum.Decimal)
                        {
                            cellStyle.Alignment = HorizontalAlignment.Right;//金额类型左对齐
                        }
                    }
                    CellStyleDic.Add(head.HeadName, cellStyle);
                }
                #endregion

            }
        }

        #endregion

        /// <summary>
        /// Excel处理
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void ExcelHandle(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            //执行处理前扩展
            ExcelHandleBefore(excelGlobalDTO);

            //设置区块
            AreaBlock<TEntity> areaBlock = new AreaBlock<TEntity>();
            areaBlock.SetAreaBlock(excelGlobalDTO);

            //设置头部颜色
            HeadColor<TEntity> headColor = new HeadColor<TEntity>();
            headColor.SetHeadColor(excelGlobalDTO);

            excelGlobalDTO.PerformanceMonitoring.Start("SetRowColor");
            //设置行颜色
            RowColor<TEntity> rowColor = new RowColor<TEntity>();
            rowColor.SetRowColor(excelGlobalDTO);
            excelGlobalDTO.PerformanceMonitoring.Stop();

            //设置锁定
            SheetLocked<TEntity> sheetLocked = new SheetLocked<TEntity>();
            sheetLocked.SetSheetLocked(excelGlobalDTO);

            //设置列隐藏
            SheetColumnHidden<TEntity> sheetColumnHidden = new SheetColumnHidden<TEntity>();
            sheetColumnHidden.SetSheetColumnHidden(excelGlobalDTO);

            //批注
            Comment<TEntity> comment = new Comment<TEntity>();

            //清空批注
            comment.ClearComment(excelGlobalDTO);

            //设置批注
            comment.SetComment(excelGlobalDTO);

            //设置列类型
            SheetColumnType<TEntity> sheetColumnType = new SheetColumnType<TEntity>();
            sheetColumnType.SetSheetColumnType(excelGlobalDTO);

            //从Excel处理
            SlaveExcel<TEntity> slaveExcel = new SlaveExcel<TEntity>();
            slaveExcel.SlaveExcelHandle(excelGlobalDTO);

            //执行处理后扩展
            ExcelHandleAfter(excelGlobalDTO);
        }

        /// <summary>
        /// Excel处理
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void ExcelHandleBefore(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            IExport<TEntity> export = ServiceContainer.GetService<IExport<TEntity>>();
            if (export != null)
            {
                export.ExcelHandleBefore(excelGlobalDTO);
            }
        }

        /// <summary>
        /// Excel处理
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        public void ExcelHandleAfter(ExcelGlobalDTO<TEntity> excelGlobalDTO)
        {
            IExport<TEntity> export = ServiceContainer.GetService<IExport<TEntity>>();
            if (export != null)
            {
                export.ExcelHandleAfter(excelGlobalDTO);
            }
        }

        /// <summary>
        /// 获取该类的最基础的类，也就是execl基础信息对象
        /// </summary>
        /// <param name="inheritClass"></param>
        /// <returns></returns>
        public Type GetExeclBaseType(Type inheritClass)
        {
            if (inheritClass == null)
            {
                return null;
            }
            if (inheritClass.Name == "ExcelRowModel")
            {
                return inheritClass;
            }
            return GetExeclBaseType(inheritClass.BaseType);
        }
    }
}

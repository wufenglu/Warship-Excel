using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Warship.Attribute;
using Warship.Attribute.Model;
using Warship.Excel.Common;
using Warship.Excel.Model;
using Warship.Excel.Model.Column;
using Warship.Attribute.Attributes;
using System.Collections;
using Warship.Excel.Import.Helper;
using static Warship.Excel.Common.ExcelHelper;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using System.Linq.Expressions;
using Warship.Utility;

namespace Warship.Excel.Import
{
    /// <summary>
    /// 导入
    /// </summary>
    public class Import<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 全局对象
        /// </summary>
        public ExcelGlobalDTO<TEntity> ExcelGlobalDTO { get; set; }

        #region 构造函数

        /// <summary>
        /// 构造函数
        /// </summary>
        public Import()
        {
            ExcelGlobalDTO = new ExcelGlobalDTO<TEntity>();
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="globalStartRowIndex">起始行</param>
        /// <param name="globalStartColumnIndex">起始列</param>
        public Import(int globalStartRowIndex, int? globalStartColumnIndex = 0)
        {
            ExcelGlobalDTO = new ExcelGlobalDTO<TEntity>();
            ExcelGlobalDTO.GlobalStartRowIndex = globalStartRowIndex;
            ExcelGlobalDTO.GlobalStartColumnIndex = globalStartColumnIndex ?? 0;
        }

        #endregion

        #region 导入

        /// <summary>
        /// 执行导入
        /// </summary>
        public virtual void Execute(string filePath)
        {
            //获取工作簿
            this.GetWorkbook(filePath);

            //获取Sheet
            this.GetSheets();

            //获取实体集合（头部、数据）对象及校验（头部、实体）
            this.GetEntityAndValidation();
        }

        /// <summary>
        /// 根据流执行
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="excelVersionEnum"></param>
        public virtual void ExecuteByStream(Stream stream, ExcelVersionEnum excelVersionEnum)
        {
            //获取工作簿
            this.GetWorkbook(stream, excelVersionEnum);

            //获取Sheet
            this.GetSheets();

            //获取实体集合（头部、数据）对象及校验（头部、实体）
            this.GetEntityAndValidation();
        }

        /// <summary>
        /// 根据流执行
        /// </summary>
        /// <param name="buffer"></param>
        /// <param name="excelVersionEnum"></param>
        public virtual void ExecuteByBuffer(byte[] buffer, ExcelVersionEnum excelVersionEnum)
        {
            using (Stream stream = new MemoryStream(buffer))
            {
                //获取工作簿
                this.GetWorkbook(stream, excelVersionEnum);
            }
            //获取Sheet
            this.GetSheets();

            //获取实体集合（头部、数据）对象及校验（头部、实体）
            this.GetEntityAndValidation();
        }

        #endregion

        #region 获取Workbook

        /// <summary>
        /// 文件路径
        /// </summary>
        /// <param name="filePath"></param>
        public void GetWorkbook(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);
            long fileLength = fileInfo.Length;
            byte[] buffers = new byte[fileLength];

            using (FileStream fs = new FileStream(filePath, FileMode.Open))
            {
                int count = fs.Read(buffers, 0, buffers.Length);
                if (count == 0)
                {
                    throw new Exception("长度为0");
                }
            }

            ExcelGlobalDTO.FileName = fileInfo.Name;
            ExcelGlobalDTO.FilePath = filePath;
            ExcelGlobalDTO.FileBytes = buffers;

            using (Stream stream = new FileStream(filePath, FileMode.Open))
            {
                ExcelGlobalDTO.ExcelVersionEnum = ExcelHelper.GetExcelVersion(fileInfo.Name);
                ExcelGlobalDTO.Workbook = ExcelHelper.GetWorkbook(stream, ExcelGlobalDTO.ExcelVersionEnum);
            }
        }

        /// <summary>
        /// 获取工作簿
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="excelVersionEnum"></param>
        public void GetWorkbook(Stream stream, ExcelVersionEnum excelVersionEnum)
        {
            ExcelGlobalDTO.PerformanceMonitoring.Start("=====【导入】获取工作簿=====");

            //获取字节
            byte[] buffers = new byte[stream.Length];
            int read;
            read = stream.Read(buffers, 0, buffers.Length);
            if (read > 0)
            {
                stream.Seek(0, SeekOrigin.Begin);
                ExcelGlobalDTO.FileBytes = buffers;
            }
            //版本、工作簿赋值
            ExcelGlobalDTO.ExcelVersionEnum = excelVersionEnum;
            ExcelGlobalDTO.Workbook = ExcelHelper.GetWorkbook(stream, excelVersionEnum);

            ExcelGlobalDTO.PerformanceMonitoring.Stop();
        }

        #endregion

        #region 获取Sheet

        /// <summary>
        /// 获取Sheet
        /// </summary>
        private void GetSheets()
        {
            ExcelGlobalDTO.PerformanceMonitoring.Start("=====【导入】获取Sheet=====");

            //如果没有设置Sheet则基于Workbook初始化Sheet，起始行起始列去全局的
            if (ExcelGlobalDTO.Sheets == null || ExcelGlobalDTO.Sheets.Count == 0)
            {
                //文件流
                Stream stream = new MemoryStream(ExcelGlobalDTO.FileBytes);
                //获取Sheet并设置
                ExcelGlobalDTO.Sheets = this.GetSheets(stream, ExcelGlobalDTO.ExcelVersionEnum);
                foreach (var item in ExcelGlobalDTO.Sheets)
                {
                    item.StartRowIndex = ExcelGlobalDTO.GlobalStartRowIndex;
                    item.StartColumnIndex = ExcelGlobalDTO.GlobalStartColumnIndex;
                }
            }
            else
            {
                //如果未设置起始行起始列，则以全局为准
                foreach (var item in ExcelGlobalDTO.Sheets)
                {
                    item.StartRowIndex = item.StartRowIndex ?? ExcelGlobalDTO.GlobalStartRowIndex;
                    item.StartColumnIndex = item.StartColumnIndex ?? ExcelGlobalDTO.GlobalStartColumnIndex;
                }
            }

            ExcelGlobalDTO.PerformanceMonitoring.Stop();
        }

        /// <summary>
        /// 获取Sheet
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="excelVersionEnum"></param>
        /// <returns></returns>
        public List<ExcelSheetModel<TEntity>> GetSheets(Stream stream, ExcelVersionEnum excelVersionEnum)
        {
            //返回结果
            List<ExcelSheetModel<TEntity>> sheets = new List<ExcelSheetModel<TEntity>>();

            IWorkbook workbook = ExcelHelper.GetWorkbook(stream, excelVersionEnum);
            //获取所有Sheet
            int sheetCount = workbook.NumberOfSheets;
            //遍历Sheet
            for (int i = 0; i < sheetCount; i++)
            {
                //不如不启用当前Sheet则跳过
                if (ExcelGlobalDTO.DisableSheetIndexs != null && ExcelGlobalDTO.DisableSheetIndexs.Contains(i))
                {
                    continue;
                }

                ISheet sheet = workbook.GetSheetAt(i);

                //获取sheet
                SheetAttribute sheetAttr = typeof(TEntity).GetCustomAttribute<SheetAttribute>() as SheetAttribute;
                if (sheetAttr != null && sheetAttr.SheetName != sheet.SheetName)
                {
                    continue;
                }

                
                //获取头部行
                IRow row = sheet.GetRow(ExcelGlobalDTO.GlobalStartRowIndex);
                if (row == null)
                {
                    continue;
                }

                //获取表头信息
                List<string> cellValues = row.Cells.Select(s => ExcelHelper.GetCellValue(s)).ToList();
                if (cellValues == null)
                {
                    continue;
                }

                //构建默认一个
                ExcelSheetModel<TEntity> sheetModel = new ExcelSheetModel<TEntity>();
                sheetModel.SheetIndex = i;
                sheetModel.SheetName = sheet.SheetName;

                //设置一个默认
                sheets.Add(sheetModel);
            }
            return sheets;
        }

        #endregion

        /// <summary>
        /// 获取实体集合（头部、数据）对象及校验（头部、实体）
        /// </summary>
        private void GetEntityAndValidation()
        {
            //判断Sheet是否存在，不存在不需要进行后续步骤
            if (ExcelGlobalDTO.Sheets == null || ExcelGlobalDTO.Sheets.Count == 0)
            {
                return;
            }

            ExcelGlobalDTO.PerformanceMonitoring.Start("=====【导入】验证头部=====");
            //验证头部
            this.ValidationHead();
            ExcelGlobalDTO.PerformanceMonitoring.Stop();

            ExcelGlobalDTO.PerformanceMonitoring.Start("=====【导入】获取头部集合=====");
            //获取头部集合
            this.GetSheetHeadList();
            ExcelGlobalDTO.PerformanceMonitoring.Stop();

            ExcelGlobalDTO.PerformanceMonitoring.Start("=====【导入】获取实体集合=====");
            //获取实体集合
            this.GetSheetEntityList();
            ExcelGlobalDTO.PerformanceMonitoring.Stop();

            ExcelGlobalDTO.PerformanceMonitoring.Start("=====【导入】校验实体=====");
            //校验实体
            this.ValidationValue();
            ExcelGlobalDTO.PerformanceMonitoring.Stop();
        }

        #region 获取头部信息

        /// <summary>
        /// 获取Sheet头部信息
        /// </summary>
        /// <returns></returns>
        public void GetSheetHeadList()
        {
            //遍历Sheet
            foreach (var sheetModel in ExcelGlobalDTO.Sheets)
            {
                //获取Sheet
                ISheet sheet = ExcelGlobalDTO.Workbook.GetSheetAt(sheetModel.SheetIndex);
                //获取头部行
                IRow headRow = sheet.GetRow(sheetModel.StartRowIndex.Value);
                if (headRow == null)
                {
                    continue;
                }

                #region 合并列：把实体属性+Excel头部列合并(二开扩展列)

                //头部实体集合：实体属性+Excel列
                List<ExcelHeadDTO> entityHeadDtoList = ExcelAttributeHelper<TEntity>.GetHeads();
                List<ExcelHeadDTO> sheetHeadDtoList = GetSheetHead(headRow.Cells);

                //获取不在实体属性中的列，并合并至头部集合
                IEnumerable<string> headNames = entityHeadDtoList.Select(s => s.HeadName);
                IEnumerable<ExcelHeadDTO> otherColumns = sheetHeadDtoList.Where(w => headNames.Contains(w.HeadName) == false);
                if (otherColumns != null)
                {
                    entityHeadDtoList.AddRange(otherColumns);
                }

                #endregion

                #region 设置列序号

                foreach (var item in entityHeadDtoList)
                {
                    //因为根据实体获取的头部对象还要很多其他的属性值，所以存在同名则以实体为主;
                    var headDto = sheetHeadDtoList.FirstOrDefault(it => it.HeadName.Equals(item.HeadName));
                    if (headDto != null)
                    {
                        item.ColumnIndex = headDto.ColumnIndex;
                    }
                }

                #endregion

                //添加头部
                sheetModel.SheetHeadList = entityHeadDtoList;
            }
        }

        /// <summary>
        /// 获取工作簿的表头
        /// </summary>
        /// <param name="cells"></param>
        protected List<ExcelHeadDTO> GetSheetHead(List<ICell> cells)
        {
            if (cells == null)
            {
                return null;
            }

            //返回
            List<ExcelHeadDTO> result = new List<ExcelHeadDTO>();

            //遍历
            foreach (ICell cell in cells)
            {
                ExcelHeadDTO headDTO = new ExcelHeadDTO
                {
                    HeadName = ExcelHelper.GetCellValue(cell),
                    ColumnIndex = cell.ColumnIndex,
                };
                result.Add(headDTO);
            }

            return result;
        }

        #endregion

        #region 实体赋值

        /// <summary>
        /// 获取Sheet实体集合
        /// </summary>
        private void GetSheetEntityList()
        {
            //结果集
            Dictionary<string, List<TEntity>> result = new Dictionary<string, List<TEntity>>();

            //遍历Sheet，目的：获取Sheet实体集合
            foreach (var sheetModel in ExcelGlobalDTO.Sheets)
            {
                //Sheet、实体集合                
                ISheet sheet = ExcelGlobalDTO.Workbook.GetSheetAt(sheetModel.SheetIndex);
                List<ExcelHeadDTO> sheetHead = sheetModel.SheetHeadList;

                //如果头部为空，则不处理当前Sheet
                if (sheetHead == null || sheetHead.Count == 0)
                {
                    continue;
                }

                //获取实体集合
                List<TEntity> entityList = GetEntityList(sheet, sheetModel);

                //添加Sheet实体集合
                sheetModel.SheetEntityList = entityList;
            }
        }

        /// <summary>
        /// 获取Sheet实体集合
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sheetModel"></param>
        private List<TEntity> GetEntityList(ISheet sheet, ExcelSheetModel<TEntity> sheetModel)
        {
            //获取图片
            List<ColumnFile> files = FileHelper.GetFiles(sheet);

            var pairs = new ConcurrentDictionary<int, TEntity>();
            var parallelOptions = MultiThreadingHelper.GetParallelOptions();
            Parallel.For(
                (sheetModel.StartRowIndex.Value) + 1,
                sheet.LastRowNum + 1,
                parallelOptions,
                rowNum =>
                {
                    TEntity entity = GetEntity(rowNum, sheetModel, sheet, files);
                    if (entity != null)
                    {
                        pairs[rowNum] = entity;
                    }
                }
                );
            return pairs.OrderBy(o => o.Key).Select(s => s.Value).ToList();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowNum"></param>
        /// <param name="sheetModel"></param>
        /// <param name="sheet"></param>
        /// <param name="files"></param>
        /// <returns></returns>
        private TEntity GetEntity(int rowNum, ExcelSheetModel<TEntity> sheetModel, ISheet sheet, List<ColumnFile> files)
        {
            //实体
            TEntity entity = new TEntity
            {
                RowNumber = rowNum
            };

            #region 行转换实体并赋值

            //获取行
            IRow row = sheet.GetRow(rowNum);

            #region 行判断及列是否有值判断
            //当前行没有任何内容                    
            if (row == null || row.Cells == null)
            {
                return null;
            }

            //判断是否有值，没有值则跳出
            List<string> cellValues = row.Cells.Select(s => ExcelHelper.GetCellValue(s)).ToList();
            if (cellValues.Exists(w => string.IsNullOrEmpty(w) == false) == false)
            {
                return null;
            }
            #endregion

            //实体类型
            Type entityType = entity.GetType();

            //遍历头部，设置属性值和其他列
            foreach (ExcelHeadDTO headDto in sheetModel.SheetHeadList)
            {
                ICell cell = row.GetCell(headDto.ColumnIndex);
                string value = ExcelHelper.GetCellValue(cell);//获取单元格的值，设置属性值

                #region 属性为空,添加其他列
                //属性为空,添加其他列
                if (string.IsNullOrEmpty(headDto.PropertyName) == true)
                {
                    EntityOtherColumnsSet(entity, sheetModel, headDto, value);
                    continue;
                }
                #endregion

                #region  属性不为空，设置属性值
                //属性不为空，设置属性值
                PropertyInfo prop = PropertyHelper.GetPropertyInfo<TEntity>(headDto.PropertyName);

                //级联Sheet
                SheetAttribute sheetAttribute = prop.GetCustomAttribute<SheetAttribute>();
                if (sheetAttribute != null)
                {
                    EntityPropertyCascadeSet(entity, prop, sheetAttribute);
                    continue;
                }

                try
                {
                    #region 非泛型列赋值

                    //列文件
                    if (prop.PropertyType == typeof(List<ColumnFile>))
                    {
                        List<ColumnFile> columnFiles = files.Where(n => n.MinRow == row.RowNum && n.MinCol == headDto.ColumnIndex).ToList();
                        if (columnFiles != null && columnFiles.Count > 0)
                        {
                            prop.SetValue(entity, columnFiles, null);
                        }
                        continue;
                    }

                    //判断值
                    if (string.IsNullOrEmpty(value))
                    {
                        continue;
                    }

                    //entity.SetPropertyValue(headDto.PropertyName, value);

                    //属性是否设置
                    bool propertyIsSetValue = this.EntityPropertySetValue(entity, prop, headDto, cell);
                    if (propertyIsSetValue == true)
                    {
                        continue;
                    }

                    //默认设置
                    prop.SetValue(entity, Convert.ChangeType(value, prop.PropertyType), null);

                    #endregion
                }
                catch (Exception ex)
                {
                    EntityPropertyErrorSet(entity, prop, headDto, ex);
                }
                #endregion
            }
            #endregion

            return entity;
        }

        #region 实体属性赋值处理

        /// <summary>
        /// 
        /// </summary>
        private object lockObj = new object();

        /// <summary>
        /// 实体其他列设置
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="sheetModel"></param>
        /// <param name="headDto"></param>
        /// <param name="value"></param>
        public virtual void EntityOtherColumnsSet(TEntity entity, ExcelSheetModel<TEntity> sheetModel, ExcelHeadDTO headDto, string value)
        {
            //其他不在属性内的列
            ColumnModel column = new ColumnModel
            {
                ColumnIndex = headDto.ColumnIndex,
                ColumnName = headDto.HeadName,
                ColumnValue = value
            };
            entity.OtherColumns.Add(column);
        }

        /// <summary>
        /// 实体属性级联设置
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="prop"></param>
        /// <param name="sheetAttribute"></param>
        private void EntityPropertyCascadeSet(TEntity entity, PropertyInfo prop, SheetAttribute sheetAttribute)
        {
            //获取Sheet信息
            ISheet propSheet = ExcelGlobalDTO.Workbook.GetSheet(sheetAttribute.SheetName);
            Type[] types = prop.PropertyType.GetGenericArguments();
            Type type = types[0];//获取泛型类型
            string key = entity.GetType().FullName + "." + prop.Name;

            #region 读取下级Sheet的内容，由于属性是遍历的因此如果一次已经获取下级Sheet，则直接从CascadeCache取
            if (CascadeCache.ChildSheetEntityList.Keys.Contains(key) == false)
            {
                lock (lockObj)
                {
                    if (CascadeCache.ChildSheetEntityList.Keys.Contains(key) == false)
                    {
                        string assemblyString = "Warship";
                        string importClassName = "Warship.Excel.Import.Import`1";
                        Type importType = Assembly.Load(assemblyString).GetType(importClassName).MakeGenericType(type);

                        //构造参数，执行构造函数
                        object importModel = System.Activator.CreateInstance(importType);
                        Type[] pt = new Type[2];
                        pt[0] = typeof(int);
                        pt[1] = typeof(int);
                        ConstructorInfo ci = importType.GetConstructor(pt);
                        ci.Invoke(new object[] { sheetAttribute.StartRowIndex, sheetAttribute.StartColumnIndex });

                        //获取非当前Sheet，全部设置为禁用，保证之取到当前的
                        int thisSheetIndex = ExcelGlobalDTO.Workbook.GetSheetIndex(sheetAttribute.SheetName);
                        List<int> disableSheetIndexs = new List<int>();
                        for (int i = 0; i < ExcelGlobalDTO.Workbook.NumberOfSheets; i++)
                        {
                            if (i != thisSheetIndex)
                            {
                                disableSheetIndexs.Add(i);
                            }
                        }
                        object excelGlobalDTO = importModel.GetType().GetProperty("ExcelGlobalDTO").GetValue(importModel);
                        excelGlobalDTO.GetType().GetProperty("DisableSheetIndexs").SetValue(excelGlobalDTO, disableSheetIndexs);

                        //执行
                        importType.GetMethod("ExecuteByBuffer").Invoke(importModel, new object[] { ExcelGlobalDTO.FileBytes, ExcelGlobalDTO.ExcelVersionEnum });
                        ExcelGlobalDTO.SlaveExcelGlobalDTO = excelGlobalDTO;

                        var globalSheets = excelGlobalDTO.GetType().GetProperty("Sheets").GetValue(excelGlobalDTO) as IEnumerable;
                        foreach (var item in globalSheets)
                        {
                            var sheetEntitys = item.GetType().GetProperty("SheetEntityList").GetValue(item) as IEnumerable;
                            CascadeCache.ChildSheetEntityList.Add(key, sheetEntitys);
                        }
                    }
                }
            }
            #endregion

            #region 级联属性赋值
            string cascadeTypeName = "Warship.Excel.Import.Helper.Cascade`2";
            Type[] cascadeTypeArguments = new Type[2];
            cascadeTypeArguments[0] = entity.GetType();
            cascadeTypeArguments[1] = type;
            Type cascadeType = Type.GetType(cascadeTypeName).MakeGenericType(cascadeTypeArguments);
            object cascadeModel = System.Activator.CreateInstance(cascadeType);
            cascadeType.GetMethod("SetEntityPropertyValues").Invoke(cascadeModel, new object[] { sheetAttribute, entity, prop, CascadeCache.ChildSheetEntityList[key] });
            #endregion
        }

        /// <summary>
        /// 实体属性赋值
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="prop"></param>
        /// <param name="headDto"></param>
        /// <param name="cell"></param>
        private bool EntityPropertySetValue(TEntity entity, PropertyInfo prop, ExcelHeadDTO headDto, ICell cell)
        {
            //变量申明
            string value = ExcelHelper.GetCellValue(cell);//获取单元格的值，设置属性值

            //日期
            if (headDto.ColumnType == Attribute.Enum.ColumnTypeEnum.Date && cell.CellType == CellType.Numeric)
            {
                //赋值、设置
                prop.SetValue(entity, cell.DateCellValue, null);
                return true;
            }

            //Decimal保留两位
            if (headDto.ColumnType == Attribute.Enum.ColumnTypeEnum.Decimal && string.IsNullOrEmpty(headDto.Format) == false)
            {
                //赋值、设置
                string newValue = decimal.Parse(value).ToString(GetNewFormatString(headDto.Format));
                prop.SetValue(entity, decimal.Parse(newValue), null);
                return true;
            }


            //GUID类型转换
            if (prop.PropertyType == typeof(Guid?))
            {
                //赋值、设置
                prop.SetValue(entity, new Guid(value), null);
                return true;
            }

            //GUID类型转换
            if (prop.PropertyType == typeof(Guid))
            {
                //赋值、设置
                prop.SetValue(entity, new Guid(value), null);
                return true;
            }
            if (prop.PropertyType == typeof(decimal) && string.IsNullOrEmpty(headDto.Format) == false)
            {
                prop.SetValue(entity, Math.Round(Convert.ToDecimal(value), int.Parse(headDto.Format), MidpointRounding.AwayFromZero), null);
                return true;
            }

            //类型是泛型类型且为空类型
            if (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                //赋值、设置
                prop.SetValue(entity, Convert.ChangeType(value, Nullable.GetUnderlyingType(prop.PropertyType)), null);
                return true;
            }

            return false;
        }

        /// <summary>
        /// 异常处理
        /// </summary>
        /// <param name="entity"></param>
        /// <param name="prop"></param>
        /// <param name="headDto"></param>
        /// <param name="ex"></param>
        private void EntityPropertyErrorSet(TEntity entity, PropertyInfo prop, ExcelHeadDTO headDto, Exception ex)
        {
            #region 异常处理

            //获取属性转换异常信息
            string convertErrorMsg = ConvertErrorAttributeHelper<TEntity>.GetPropertyConvertErrorInfo(prop.Name);
            if (string.IsNullOrEmpty(convertErrorMsg) == true)
            {
                convertErrorMsg = ex.Message;
            }

            //设置异常信息
            ColumnErrorMessage errorMsg = new ColumnErrorMessage
            {
                PropertyName = prop.Name,
                ErrorMessage = convertErrorMsg,
                ColumnName = headDto.HeadName
            };
            entity.ColumnErrorMessage.Add(errorMsg);
            #endregion
        }

        #endregion

        #endregion

        #region 校验

        /// <summary>
        /// 校验头部
        /// </summary>
        public virtual void ValidationHead()
        {
            //遍历Sheet进行头部校验
            foreach (var sheetModel in ExcelGlobalDTO.Sheets)
            {
                ISheet sheet = ExcelGlobalDTO.Workbook.GetSheetAt(sheetModel.SheetIndex);

                //获取头部行
                IRow row = sheet.GetRow(sheetModel.StartRowIndex.Value);
                if (row == null)
                {
                    throw new Exception(ExcelGlobalDTO.ExcelValidationMessage.Clgyl_Common_Import_TempletError);
                }

                //获取表头信息
                List<string> cellValues = row.Cells.Select(s => ExcelHelper.GetCellValue(s)).ToList();

                List<ExcelHeadDTO> headDtoList = ExcelAttributeHelper<TEntity>.GetHeads();
                //头部校验
                foreach (ExcelHeadDTO dto in headDtoList)
                {
                    //校验必填的，判断表头是否在excel中存在
                    if (dto.IsValidationHead == true && cellValues.Contains(dto.HeadName) == false)
                    {
                        throw new Exception(ExcelGlobalDTO.ExcelValidationMessage.Clgyl_Common_Import_TempletError);
                    }
                }
            }

            //动态列验证
            DynamicColumn<TEntity>.ValidationHead(ExcelGlobalDTO, ValidationModelEnum.DynamicColumn);

            //头部校验后处理，用于特殊处理
            ValidationHeaderAfter();
        }

        /// <summary>
        /// 执行
        /// </summary>
        /// <returns></returns>
        public virtual void ValidationHeaderAfter()
        {
            IImport<TEntity> import = ServiceContainer.GetService<IImport<TEntity>>();
            if (import != null)
            {
                import.ValidationHeaderAfter(ExcelGlobalDTO);
            }
        }

        /// <summary>
        /// 验证内容
        /// </summary>
        public virtual void ValidationValue()
        {
            //遍历Sheet实体集合
            foreach (var sheet in ExcelGlobalDTO.Sheets)
            {
                //获取Sheet头部实体集合
                var headDtoList = sheet.SheetHeadList;

                //为空判断
                if (sheet.SheetEntityList == null)
                {
                    continue;
                }

                var pairs = new ConcurrentDictionary<int, List<ValidationResult>>();
                var parallelOptions = MultiThreadingHelper.GetParallelOptions();
                Parallel.ForEach(sheet.SheetEntityList, parallelOptions, entity =>
                {
                    pairs.TryAdd(entity.RowNumber, ValidationHelper.Exec<TEntity>(entity));
                });

                foreach (var item in sheet.SheetEntityList)
                {
                    #region 选项值校验
                    //校验文本框的值是否在下拉框选项中
                    foreach (var headDto in headDtoList)
                    {
                        #region 基础判断
                        //判断是否为选项列
                        if (headDto.ColumnType != Attribute.Enum.ColumnTypeEnum.Option)
                        {
                            continue;
                        }
                        //判断选项是否存在
                        if (sheet.ColumnOptions == null)
                        {
                            continue;
                        }

                        //判断属性是否存在
                        if (string.IsNullOrEmpty(headDto.PropertyName))
                        {
                            continue;
                        }
                        #endregion

                        //其他列的键值
                        string key = null;

                        #region 获其他列的键值
                        if (sheet.ColumnOptions.Keys.Contains(headDto.HeadName))
                        {
                            key = headDto.HeadName;
                        }
                        if (sheet.ColumnOptions.Keys.Contains(headDto.PropertyName))
                        {
                            key = headDto.PropertyName;
                        }
                        //判断键是否为空
                        if (key == null)
                        {
                            continue;
                        }
                        #endregion

                        //变量设置
                        PropertyInfo propertyInfo = item.GetType().GetProperty(headDto.PropertyName);
                        string value = propertyInfo.GetValue(item).ToString();

                        #region 校验
                        //类型判断decimal
                        if (propertyInfo.PropertyType == typeof(decimal) || propertyInfo.PropertyType == typeof(decimal?))
                        {
                            List<decimal> options = sheet.ColumnOptions[key].Select(n => Convert.ToDecimal(n)).ToList();
                            if (options.Contains(Convert.ToDecimal(value)) == true)
                            {
                                continue;
                            }
                        }

                        //类型判断double
                        if (propertyInfo.PropertyType == typeof(double) || propertyInfo.PropertyType == typeof(double?))
                        {
                            List<double> options = sheet.ColumnOptions[key].Select(n => Convert.ToDouble(n)).ToList();
                            if (options.Contains(Convert.ToDouble(value)) == true)
                            {
                                continue;
                            }
                        }

                        //判断输入的值是否在选项值中                        
                        if (sheet.ColumnOptions[key].Contains(value) == true)
                        {
                            continue;
                        }

                        #endregion

                        #region 不通过则提示异常
                        //异常信息
                        ColumnErrorMessage errorMsg = new ColumnErrorMessage
                        {
                            PropertyName = headDto.PropertyName,
                            ColumnName = headDto.HeadName,
                            ErrorMessage = ExcelGlobalDTO.ExcelValidationMessage.Clgyl_Common_Import_NotExistOptions
                        };
                        item.ColumnErrorMessage.Add(errorMsg);
                        #endregion
                    }
                    #endregion
                    
                    #region  实体特性验证
                    List<ValidationResult> result = pairs[item.RowNumber];
                    if (result == null)
                    {
                        continue;
                    }
                    foreach (var msg in result)
                    {
                        //异常信息
                        ColumnErrorMessage errorMsg = new ColumnErrorMessage
                        {
                            PropertyName = msg.PropertyName,
                            ErrorMessage = msg.ErrorMessage
                        };

                        //设置列信息
                        var headDto = headDtoList.Where(w => w.PropertyName == msg.PropertyName).FirstOrDefault();
                        if (headDto != null)
                        {
                            errorMsg.ColumnName = headDto.HeadName;
                        }

                        //添加至集合
                        item.ColumnErrorMessage.Add(errorMsg);
                    }
                    #endregion
                }
            }

            //动态列验证
            DynamicColumn<TEntity>.ValidationValue(ExcelGlobalDTO, ValidationModelEnum.DynamicColumn);

            //验证值后
            this.ValidationValueAfter();
        }

        /// <summary>
        /// 验证数据
        /// </summary>
        public virtual void ValidationValueAfter()
        {
            IImport<TEntity> import = ServiceContainer.GetService<IImport<TEntity>>();
            if (import != null)
            {
                import.ValidationValueAfter(ExcelGlobalDTO);
            }
        }

        #endregion
    }
}


## 一、1.	Warship定位
Warship是一款基于NPOI的优秀的Excel导入导出组件，基于实体、特性、注入、多线程实现Excel的导入和导出，是一款使用方便、扩展性强、性能优良的组件，开发者不需要关注如何操作NPOI，只需要对实体进行操作即可实现导入导出，使用它可以低成本、高质量的实现业务场景的导入和导出。

QQ群（1154777006）


## 二、Warship背景
面向各业务系统，在业务使用过程中对于清单化的数据往往都有强烈的Excel导入导出需求，以避免通过系统逐条录入，利用Excel的便利性快速的完成数据的维护和导入。但目前市面上能够提供的组件也就只有NPIO等基础的组件，这些组件更多聚焦于Excel交互，对于业务开发者的支撑非常少，开发者需要编写大量的基础代码来实现业务逻辑，因此急需一个组件能够支持业务代码的快速开发，接管大量基础的功能，让开发者把经理聚焦在高价值的业务兑现上。

## 三、Warship解决的痛点
开发人员需要基于每一个场景都实现一次针对NPOI的操作来实现导入导出，业务场景中往往不同的场景导入导出的要求还不一样，有的有特殊的业务逻辑校验、有的有动态列、有的有多Sheet级联，往往开发人员没有很好的封装思路，导致每个场景都实现一套导入导出逻辑，代码各种交织，调用非常混乱，后续改动一个基础场景各种影响点。特别是对于Excel导入有很多场景是通用，比如头部校验、必填、长度、范围、格式校验等，每一个场景开发者都要单独实现一次。即耗时又容易遗漏，随着场景的增多，代码扩展性、质量和危害性都面临巨大的问题。

## 四、Warship代码简介

1）标准导入导出指基于标准组件即可完成导入导出功能，不需要进行扩展开发。同时导入导出都是实体化的，可以通过实体操作来进行Excel的操作。

#### 2）特性介绍

&emsp;2.1）ExcelHead：ExcelHead为属性对应Excel的单元格头部，通过该特性可以锁定Excel里面的单元格进行属性值设置，同时对Excel进行锁定、隐藏、头部颜色、整列颜色、列类型（文本、选项、日期、金额等）、格式设置

&emsp;2.2）Required：添加Required特性即为必填校验，可设置校验不通过时的提示信息ErrorMessage
  
&emsp;2.3）Length：添加Length特性即进行长度校验，可设置校验不通过时的提示信息ErrorMessage
  
&emsp;2.4）Range：添加Range特性即进行范围控制，可设置校验不通过时的提示信息ErrorMessage

&emsp;2.5）Format：添加Format特性即限制字段的输入格式，可设置校验不通过时的提示信息ErrorMessage，格式校验内置4个标准：邮箱、电话、移动电话、身份证，如果内置的不够，可通过正则进行设置，使用特性的重载函数即可

#### 3）ExcelGlobalDTO介绍

&emsp;3.1）ExcelGlobalDTO为Excel级别的全局信息，包含Excel的文件信息、所有Sheet起始行、所有Sheet起始列、Sheet实体集合

&emsp;3.2）可通过ExcelGlobalDTO设置跟Excel的Sheet相关的设置，如禁用Sheet、设置活动的Sheet

#### 4）ExcelSheetModel介绍

&emsp;4.1）ExcelSheetModel为Sheet级别的信息，包含Sheet的名称（SheetName）、序号（SheetIndex）、起始行（StartRowIndex）、起始列（StartColumnIndex）、说明（AreaBlock）

&emsp;4.2）导入Excel后，可获取到Sheet头部集合SheetHeadList、Sheet实体集合SheetEntityList

&emsp;4.3）可设置Sheet内列的选项值，通过ColumnOptions进行单元格输入的限制

#### 5）ExcelRowModel介绍

&emsp;5.1）所有导入导出DTO对象必须继承ExcelRowModel，ExcelRowModel为组件级的封装

&emsp;5.2）ExcelRowModel为Excel行对象，里面包含所有行相关的信息，如行号、实体中未定义列的单元格信息集合

&emsp;5.3）可以通过实体对ExcelRowModel中的属性进行样式设置（RowStyleSet）、是否删除行设置（IsDeleteRow）、动态列设置（SetDynamicColumns）

&emsp;5.3）当调用组件导入后，可以通过ExcelRowModel的ColumnErrorMessage获取到本行内单元格验证不通过信息，同时也可以基于组件以外的业务校验向ColumnErrorMessage追加异常信息，导出的时候会向单元格上打批注。

&emsp;5.4）也可通过ExcelGlobalDTO的GetColumnErrorMessages获取Excel的所有验证不通过信息。

## 二、代码示例

> ### 标准导入导出-代码示例

```csharp
/// <summary>
/// 合同
/// </summary>
[Serializable]
public class ContractImportDTO : ExcelRowModel
{

    /// <summary>
    /// 合同名称
    /// </summary>
    [ExcelHead("合同名称", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
    [Required(ErrorMessage = "合同名称必填")]
    [Length(100, ErrorMessage = "长度不能超过100")]
    public string Name { get; set; }

    /// <summary>
    /// 合同编码
    /// </summary>
    [ExcelHead("合同编码", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
    [Required(ErrorMessage = "合同名称必填")]
    [Length(100, ErrorMessage = "长度不能超过100")]
    public string Code { get; set; }

    /// <summary>
    /// 合同编码
    /// </summary>
    [ExcelHead("甲方单位", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
    [Required(ErrorMessage = "合同名称必填")]
    [Length(100, ErrorMessage = "长度不能超过100")]
    public string JfProvider { get; set; }

    /// <summary>
    /// 合同编码
    /// </summary>
    [ExcelHead("乙方单位", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
    [Required(ErrorMessage = "合同名称必填")]
    [Length(100, ErrorMessage = "长度不能超过100")]
    public string YfProvider { get; set; }

    /// <summary>
    /// 合同编码
    /// </summary>
    [ExcelHead("合同金额", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
    [Range(0, 999999999, ErrorMessage = "超出范围")]
    public decimal? Amount { get; set; }
    
    [ExcelHead("邮箱")]
    [Format(FormatEnum.Email,ErrorMessage ="格式错误")]
    public string Email { get; set; }
}

public class 标准导入导出
{
    public void Execute()
    {
        string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\动态添加特性启用禁用.xlsx";

        //导入
        Import<Mysoft.Clgyl.Demo.DemoDTO.ContractImportDTO> import = new Import<Mysoft.Clgyl.Demo.DemoDTO.ContractImportDTO>(1);
        import.Execute(excelPath);

        Export<Mysoft.Clgyl.Demo.DemoDTO.ContractImportDTO> export = new Export<Mysoft.Clgyl.Demo.DemoDTO.ContractImportDTO>();
        export.ExportMemoryStream(import.ExcelGlobalDTO);
    }
}                                                
```


> ### 材料图片导入-代码示例
```csharp
    public class 材料图片导入
    {
        public void Execute()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\图片导入.xlsx";

            //导入
            Import<Warship.Demo.DemoDTO.ProductDTO> import = new Import<Warship.Demo.DemoDTO.ProductDTO>();
            import.ExcelGlobalDTO.DisableSheetIndexs = new List<int> { 1,2,3,4,5 };
            import.Execute(excelPath);

            Export<Warship.Demo.DemoDTO.ProductDTO> export = new Export<Warship.Demo.DemoDTO.ProductDTO>();
            export.Execute(import.ExcelGlobalDTO);
        }
    }
```


> ### 动态添加特性启用禁用-代码示例
```csharp
class 动态添加特性启用禁用
    {
        /// <summary>
        /// 执行
        /// </summary>
        public void Execute()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\动态添加特性启用禁用.xls";

            //设置属性特性
            Dictionary<string, BaseAttribute> dic = new Dictionary<string, BaseAttribute>
            {
                { "Name", new RequiredAttribute { ErrorMessage = "名称必填***" } },
                { "Code", new RequiredAttribute { ErrorMessage = "编码必填***" } }
            };

            //TODO执行启用特性:待优化（统一设置特性，不用单独一个个特性设置）
            AttributeFactory<DemoDTO.ContractImportDTO>.GetValication(ValidationTypeEnum.Required).EnableAttributes(dic);

            //执行禁用特性
            AttributeFactory<DemoDTO.ContractImportDTO>.GetValication(ValidationTypeEnum.Required).DisableAttributes(new List<string> {
                "Name","Code"
            });

            //导入
            Import<DemoDTO.ContractImportDTO> import = new Import<DemoDTO.ContractImportDTO>(1);
            import.Execute(excelPath);

            //导出
            Export<DemoDTO.ContractImportDTO> export = new Export<DemoDTO.ContractImportDTO>();
            export.Execute(import.ExcelGlobalDTO);
        }
    }
```

> ### 二开扩展增加导出列-代码示例
```csharp
class 二开扩展增加导出列
    {
        /// <summary>
        /// 执行
        /// </summary>
        public void Execute()
        {
            string excelDynamicPath = Directory.GetCurrentDirectory() + "\\..\\Template\\动态列.xls";
            new DynamicColumnDTO().SetDynamicColumns(new List<ColumnModel> {
                new ColumnModel{
                    ColumnName = "动态列A",
                    ColumnValidations=new List<ColumnValidationModel>{
                        new ColumnValidationModel()
                        {
                            ValidationTypeEnum = ValidationTypeEnum.Required,
                            RequiredAttribute = new RequiredAttribute() { ErrorMessage = "动态列A必填" }
                        }
                    }
                },
                new ColumnModel{
                    ColumnName = "动态列B",
                    ColumnValidations=new List<ColumnValidationModel>{
                        new ColumnValidationModel()
                        {
                            ValidationTypeEnum = ValidationTypeEnum.Required,
                            RequiredAttribute = new RequiredAttribute() { ErrorMessage = "动态列B必填" }
                        }
                    }
                }
            });

            Import<DynamicColumnDTO> import = new Import<DynamicColumnDTO>();
            import.ExcelGlobalDTO.SetDefaultSheet("合同材料");
            import.Execute(excelDynamicPath);

            Export<DynamicColumnDTO> export = new Export<DynamicColumnDTO>();
            export.Execute(import.ExcelGlobalDTO);


        }
    }
```

> ### 导入注入-代码示例
```csharp
/// <summary>
    /// 导入注入
    /// </summary>
    public class ImportInjection : IImport<DemoDTO.ContractImportDTO>
    {
        public void ValidationHeaderAfter(ExcelGlobalDTO<DemoDTO.ContractImportDTO> ExcelGlobalDTO)
        {
            return;
        }

        public void ValidationValueAfter(ExcelGlobalDTO<DemoDTO.ContractImportDTO> ExcelGlobalDTO)
        {
            return;
        }
    }

    /// <summary>
    /// 导出注入
    /// </summary>
    public class ExportInjection : IExport<DemoDTO.ContractImportDTO>
    {
        public void ExcelHandleAfter(ExcelGlobalDTO<DemoDTO.ContractImportDTO> excelGlobalDTO)
        {
            return;
        }

        public void ExcelHandleBefore(ExcelGlobalDTO<DemoDTO.ContractImportDTO> excelGlobalDTO)
        {
            return;
        }
    }

    class 二开扩展注入
    {
        /// <summary>
        /// 执行
        /// </summary>
        public void Execute()
        {
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\合同导入性能测试.xls";

            //导入注入
            ServiceContainer.Register<IImport<DemoDTO.ContractImportDTO>, ImportInjection>();            

            //执行导入
            Import<DemoDTO.ContractImportDTO> import = new Import<DemoDTO.ContractImportDTO>(0);
            import.ExcelGlobalDTO.SetDefaultSheet();
            import.Execute(excelPath);

            //导出注入
            ServiceContainer.Register<IExport<DemoDTO.ContractImportDTO>, ExportInjection>();

            //执行导出
            Export<DemoDTO.ContractImportDTO> export = new Export<DemoDTO.ContractImportDTO>();
            export.Execute(import.ExcelGlobalDTO);
        }
    }
```

> ### 级联-代码示例
```csharp
/// <summary>
    /// 合同
    /// </summary>
    [Serializable]
    public class ContractImportDTO : ExcelRowModel
    {

        /// <summary>
        /// 合同名称
        /// </summary>
        [ExcelHead("合同名称", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        [Required(ErrorMessage = "合同名称必填")]
        [Length(100, ErrorMessage = "长度不能超过100")]
        public string Name { get; set; }

        /// <summary>
        /// 合同编码
        /// </summary>
        [ExcelHead("合同编码", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        [Required(ErrorMessage = "合同编码必填")]
        [Length(100, ErrorMessage = "长度不能超过100")]
        public string Code { get; set; }

        /// <summary>
        /// 材料
        /// </summary>
        [Sheet(SheetName = "合同材料",MasterEntityProperty = "Code",SlaveEntityProperty = "ContractCode")]
        [ExcelHead("合同材料", IsLocked = false, IsHiddenColumn = false, ColumnWidth = 8)]
        public List<ContractProductImportDTO> Products { get; set; }

        public decimal? HtAmount { get; set; }
    }
    
	public class 级联
    {
        /// <summary>
        /// 执行
        /// </summary>
        public void Execute()
        {
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            string dir = System.Environment.CurrentDirectory;

            //级联处理
            string excelPath = Directory.GetCurrentDirectory() + "\\..\\Template\\合同导入模板-级联.xls";
            Import<ContractImportDTO> import = new Import<ContractImportDTO>(1);
            import.ExcelGlobalDTO.SetDefaultSheet();
            import.Execute(excelPath);

            Export<ContractImportDTO> export = new Export<ContractImportDTO>();
            export.Execute(import.ExcelGlobalDTO);

            var errors = import.ExcelGlobalDTO.GetColumnErrorMessages();

            return;
        }
    }
```
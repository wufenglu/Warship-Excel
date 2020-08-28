using System;
using System.CodeDom.Compiler;
using System.Reflection;
using System.Text;
using Warship.Excel.Model;
using Warship.Excel.Model.Column;
using System.Linq;

namespace Warship.Utility
{
    /// <summary>
    /// 实体代理
    /// </summary>
    public class EntityProxy
    {

        /// <summary>
        /// 生成代理
        /// </summary>
        /// <param name="excelGlobalDTO"></param>
        /// <returns></returns>
        public static object GenerateProxy<TEntity>(ExcelGlobalDTO<TEntity> excelGlobalDTO) where TEntity : ExcelRowModel
        {
            string classCode = @"
                using System; 
                using System.Collections.Generic; 
                namespace Warship.Utility.CodeCompiler {{
                    public class ProxyEntity{{
                        {0}
                    }}
                }}
            ";

            //创建代码串
            StringBuilder attrCode = new StringBuilder();
            PropertyInfo[] propertyInfos = typeof(TEntity).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in propertyInfos) {
                if (prop.PropertyType.IsValueType)
                {
                    attrCode.AppendFormat("public {0}  {1} {{ set; get; }}", prop.PropertyType.Name, prop.Name);
                    attrCode.AppendLine();
                }
            }
            var dynamicColumn = Activator.CreateInstance<TEntity>().GetDynamicColumns();
            if (dynamicColumn != null)
            {
                foreach (ColumnModel prop in dynamicColumn)
                {
                    attrCode.AppendFormat("public string  {0} {{ set; get; }}", prop.PropertyName);
                    attrCode.AppendLine();
                }
            }
            if (excelGlobalDTO.Sheet.ColumnConfig != null)
            {
                foreach (ColumnModel prop in excelGlobalDTO.Sheet.ColumnConfig)
                {
                    attrCode.AppendFormat("public string  {0} {{ set; get; }}", prop.PropertyName);
                    attrCode.AppendLine();
                }
            }

            string compilerCode = string.Format(classCode, attrCode.ToString());

            //编译器的传入参数
            CompilerParameters cp = new CompilerParameters();
            cp.ReferencedAssemblies.Add("system.dll"); //添加程序集 system.dll 的引用
            cp.GenerateExecutable = false; //不生成可执行文件
            cp.GenerateInMemory = true; //在内存中运行

            //创建C#编译器实例
            CodeDomProvider provider = CodeDomProvider.CreateProvider("C#");

            //得到编译器实例的返回结果
            CompilerResults cr = provider.CompileAssemblyFromSource(cp, compilerCode);

            //如果有错误则抛异常
            if (cr.Errors.HasErrors)
            {
                StringBuilder error = new StringBuilder(); //创建错误信息字符串
                foreach (CompilerError err in cr.Errors) //遍历每一个出现的编译错误
                {
                    error.Append(err.ErrorText); //添加进错误文本，每个错误后换行
                    error.Append(Environment.NewLine);
                }
            }

            ////获取编译器实例的程序集
            Assembly assembly = cr.CompiledAssembly;

            //获取类型
            Type type = assembly.GetType("Warship.Utility.CodeCompiler.ProxyEntity");
            
            return type;
        }
    }
}





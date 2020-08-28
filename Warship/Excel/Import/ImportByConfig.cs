using Warship.Excel.Model;
using Warship.Excel.Import.Helper;
using System;
using Warship.Attribute.Model;
using System.Reflection;
using System.Xml;
using System.Linq;
using Warship.Utility;
using System.IO;

namespace Warship.Excel.Import
{
    /// <summary>
    /// 导入
    /// </summary>
    public class ImportByConfig<TEntity>  : Import<TEntity> where TEntity : ExcelRowModel, new()
    {
        /// <summary>
        /// 构造函数
        /// </summary>
        public ImportByConfig() { }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="globalStartRowIndex">起始行</param>
        /// <param name="globalStartColumnIndex">起始列</param>
        public ImportByConfig(int globalStartRowIndex, int globalStartColumnIndex = 0) : base(globalStartRowIndex, globalStartColumnIndex)
        {
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="xmlPath">xml路径</param>
        /// <param name="globalStartRowIndex">起始行</param>
        /// <param name="globalStartColumnIndex">起始列</param>
        public ImportByConfig(string xmlPath, int globalStartRowIndex = 0, int globalStartColumnIndex = 0) : base(globalStartRowIndex, globalStartColumnIndex)
        {
            ExcelSheetModel<TEntity> sheetModel = new ExcelSheetModel<TEntity>();
            ExcelGlobalDTO.Sheets.Add(sheetModel);

            XmlDocument xmlDoc = new XmlDocument();

            //判断二开文件
            FileInfo fileInfo = new FileInfo(xmlPath);
            string fileName = fileInfo.Name.Replace(fileInfo.Extension, ".custom" + fileInfo.Extension);
            string customPath = fileInfo.DirectoryName + "/" + fileName;
            if (File.Exists(customPath))
            {
                xmlDoc.Load(customPath);
            }
            else
            {
                xmlDoc.Load(xmlPath);
            }

            XmlNodeList xmlNodes = xmlDoc.SelectSingleNode("/Excel/Sheets").ChildNodes;
            foreach (XmlNode sheet in xmlNodes) //Sheet
            {
                foreach (XmlNode xmlNode in sheet.ChildNodes) //Node
                {
                    switch (xmlNode.Name)
                    {
                        case "AreaBlock":
                            XmlElement areaBlockEle = xmlNode as XmlElement;
                            ExcelGlobalDTO.Sheet.AreaBlock = ConfigHelper.GetAreaBlock(areaBlockEle);
                            break;
                        case "Header":
                            sheetModel.ColumnConfig = ConfigHelper.GetHeader(xmlNode);
                            break;
                    }
                }
            }            
        }

        /// <summary>
        /// 校验头部
        /// </summary>
        public override void ValidationHead()
        {
            //基类实体校验
            base.ValidationHead();

            //动态列验证
            DynamicColumn<TEntity>.ValidationHead(ExcelGlobalDTO);

            //头部校验后处理，用于特殊处理
            ValidationHeaderAfter();
        }

        /// <summary>
        /// 验证内容
        /// </summary>
        public override void ValidationValue()
        {
            //调用基类实体验证
            base.ValidationValue();

            //动态列验证
            DynamicColumn<TEntity>.ValidationValue(ExcelGlobalDTO);

            //验证值后
            this.ValidationValueAfter();
        }
    }
}

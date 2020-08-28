using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Warship.Attribute.Attributes;
using Warship.Attribute.Enum;
using Warship.Excel.Model;
using Warship.Excel.Model.Column;

namespace Warship.Utility
{
    /// <summary>
    /// 配置帮助类
    /// </summary>
    public class ConfigHelper
    {

        /// <summary>
        /// 获取区块
        /// </summary>
        /// <param name="areaBlock"></param>
        /// <returns></returns>
        public static AreaBlock GetAreaBlock(XmlElement areaBlock)
        {
            if (areaBlock == null)
            {
                return null;
            }
            AreaBlock entity = new AreaBlock
            {
                StartRowIndex = areaBlock.Attributes["StartRowIndex"]?.Value.ToInt() ?? 0,
                EndRowIndex = areaBlock.Attributes["EndRowIndex"]?.Value.ToInt() ?? 0,
                StartColumnIndex = areaBlock.Attributes["RowColumnIndex"]?.Value.ToInt() ?? 0,
                EndColumnIndex = areaBlock.Attributes["EndColumnIndex"]?.Value.ToInt() ?? 0,
                Height = areaBlock.Attributes["Height"]?.Value.ToShort() ?? 0,
                Content = areaBlock.Attributes["Content"]?.Value,
                LanguageKey = areaBlock.Attributes["LanguageKey"]?.Value
            };
            return entity;
        }

        /// <summary>
        /// 获取头部
        /// </summary>
        /// <param name="header"></param>
        /// <returns></returns>
        public static List<ColumnModel> GetHeader(XmlNode header)
        {
            if (header == null)
            {
                return null;
            }

            //返回配置集合
            List<ColumnModel> entitys = new List<ColumnModel>();

            //获取列配置信息
            XmlNodeList xmlNodes = header.SelectSingleNode("Columns").ChildNodes;
            foreach (var xmlNode in xmlNodes)
            {                
                XmlElement columnEle = xmlNode as XmlElement;
                ColumnModel entity = new ColumnModel
                {
                    PropertyName = columnEle.Attributes["Field"]?.Value.ToStr(),
                    ColumnName = columnEle.Attributes["Name"]?.Value.ToStr(),
                    LanguageKey = columnEle.Attributes["LanguageKey"]?.Value.ToStr(),
                    ColumnType = GetColumnTypeEnum(columnEle.Attributes["Type"]?.Value)
                };

                //错误信息
                XmlElement convertErrorEle = ((XmlNode)xmlNode).SelectSingleNode("ConvertError") as XmlElement;
                if (convertErrorEle != null)
                {
                    entity.ColumnConvertError = new ColumnConvertErrorModel
                    {
                        ErrorMessage = convertErrorEle.Attributes["ErrorMessage"]?.Value,
                        LanguageKey = convertErrorEle.Attributes["LanguageKey"]?.Value,
                        LanguageAssembly = convertErrorEle.Attributes["LanguageAssembly"]?.Value
                    };
                }

                //验证信息
                XmlNodeList validationNodes = ((XmlNode)xmlNode).SelectSingleNode("Validation").ChildNodes;
                foreach (var validationNode in validationNodes)
                {
                    XmlElement validationEle = validationNode as XmlElement;
                    ColumnValidationModel columnValidationModel = new ColumnValidationModel();
                    switch (validationEle.Name)
                    {
                        case "Required":
                            columnValidationModel.ValidationTypeEnum = ValidationTypeEnum.Required;
                            columnValidationModel.RequiredAttribute = new RequiredAttribute
                            {
                                ErrorMessage = validationEle.Attributes["ErrorMessage"]?.Value,
                                LanguageKey = validationEle.Attributes["LanguageKey"]?.Value,
                                LanguageAssembly = validationEle.Attributes["LanguageAssembly"]?.Value
                            };
                            break;
                        case "Length":
                            columnValidationModel.ValidationTypeEnum = ValidationTypeEnum.Length;
                            columnValidationModel.LengthAttribute = new LengthAttribute
                            {
                                ErrorMessage = validationEle.Attributes["ErrorMessage"]?.Value,
                                LanguageKey = validationEle.Attributes["LanguageKey"]?.Value,
                                Length = validationEle.Attributes["Length"]?.Value.ToInt() ?? 0,
                                LanguageAssembly = validationEle.Attributes["LanguageAssembly"]?.Value
                            };
                            break;
                        case "Format":
                            columnValidationModel.ValidationTypeEnum = ValidationTypeEnum.Format;
                            columnValidationModel.FormatAttribute = new FormatAttribute
                            {
                                ErrorMessage = validationEle.Attributes["ErrorMessage"]?.Value,
                                LanguageKey = validationEle.Attributes["LanguageKey"]?.Value,
                                Regex = validationEle.Attributes["Regex"]?.Value,
                                FormatEnum = FormatEnum.Email,//TODO
                                LanguageAssembly = validationEle.Attributes["LanguageAssembly"]?.Value
                            };
                            break;
                        case "Range":
                            columnValidationModel.ValidationTypeEnum = ValidationTypeEnum.Range;
                            columnValidationModel.RangeAttribute = new RangeAttribute()
                            {
                                ErrorMessage = validationEle.Attributes["ErrorMessage"]?.Value,
                                LanguageKey = validationEle.Attributes["LanguageKey"]?.Value,
                                Minimum = validationEle.Attributes["Minimum"]?.Value,
                                Maximum = validationEle.Attributes["Maximum"]?.Value,
                                LanguageAssembly = validationEle.Attributes["LanguageAssembly"]?.Value
                            };
                            break;
                    }
                    entity.ColumnValidations.Add(columnValidationModel);
                }

                entitys.Add(entity);
            }
            return entitys;
        }

        /// <summary>
        /// 获取列类型枚举
        /// </summary>
        /// <param name="columnType"></param>
        /// <returns></returns>
        public static ColumnTypeEnum GetColumnTypeEnum(string columnType)
        {
            Array enumValues = Enum.GetValues(typeof(ColumnTypeEnum));
            foreach (ColumnTypeEnum enumValue in enumValues)
            {
                if (enumValue.ToStr() == columnType)
                {
                    return enumValue;
                }
            }
            return ColumnTypeEnum.Text;
        }
    }
}

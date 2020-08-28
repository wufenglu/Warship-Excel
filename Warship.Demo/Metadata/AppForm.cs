using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Warship.Demo.Metadata
{
    public class ColumnInfo {
        public string Title { get; set; }
        public string Field { get; set; }
        public string DataType { get; set; }
        public string Required { get; set; }
        public string MaxLength { get; set; }
        public string DefaultValue { get; set; }
        public string ControlType { get; set; }
        public ComboBox ComboBox { get; set; }
    }
    public class ComboBox
    {
        public string Text { get; set; }
        public string Value { get; set; }
    }

    public class AppForm
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="metadataPath"></param>
        public void GetAppFormColumnInfo(string metadataPath) {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, metadataPath);

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(metadataPath);
            XmlNodeList nodeList = xmlDoc.SelectSingleNode("column").ChildNodes;
        }
    }
}

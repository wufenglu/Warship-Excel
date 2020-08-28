using Warship.Excel.Model;
using Warship.Excel.Model.Column;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Common
{
    /// <summary>
    /// 获取文件
    /// </summary>
    public static class FileHelper
    {
        /// <summary>
        /// 获取文件
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>

        public static List<ColumnFile> GetFiles(ISheet sheet)
        {
            return PictureHelper.GetAllPictureInfos(sheet, new WorkSpace());
        }

        /// <summary>
        /// 获取站点下面的文件流，
        /// 路径Demo：Clgyl/Template/合同材料导入模板.xls
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static FileStream GetFileStream(string filePath)
        {
            string fullPath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, filePath);
            //文件是否存在
            if (File.Exists(fullPath) == false)
            {
                return null;
            }
            return File.OpenRead(fullPath);
        }
    }
}

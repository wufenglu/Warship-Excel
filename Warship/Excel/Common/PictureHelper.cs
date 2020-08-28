using Warship.Excel.Model;
using Warship.Excel.Model.Column;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warship.Excel.Common
{
    /// <summary>
    /// 图片帮助类
    /// </summary>
    public static class PictureHelper
    {
        /// <summary>
        /// 获取所有图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="workSpace"></param>
        /// <returns></returns>
        public static List<ColumnFile> GetAllPictureInfos(ISheet sheet, WorkSpace workSpace)
        {
            if (sheet is HSSFSheet)//2003版本
            {
                return GetAllPictureInfos((HSSFSheet)sheet, workSpace);
            }
            else if (sheet is XSSFSheet)//2007版本
            {
                return GetAllPictureInfos((XSSFSheet)sheet, workSpace);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 获取范围内的所有图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="workSpace"></param>
        /// <returns></returns>
        private static List<ColumnFile> GetAllPictureInfos(HSSFSheet sheet, WorkSpace workSpace)
        {
            //execl中的图片信息
            List<ColumnFile> fileInfoList = new List<ColumnFile>();
            //获取工作表中的
            var shapeContainer = sheet.DrawingPatriarch as HSSFShapeContainer;
            if (null != shapeContainer)
            {
                //获取子集
                var shapeList = shapeContainer.Children;
                foreach (var shape in shapeList)
                {
                    if (shape is HSSFPicture)
                    {
                        //获取图片对象
                        var picture = (HSSFPicture)shape;
                        var anchor = (HSSFClientAnchor)picture.Anchor;
                        //判断位置是否在需要范围内
                        if (IsInternalOrIntersect(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2, workSpace))
                        {
                            ColumnFile entity = new ColumnFile()
                            {
                                //图片所在开始行
                                MinRow = anchor.Row1,
                                //图片所在结束行
                                MaxRow = anchor.Row2,
                                // 图片所在开始列
                                MinCol = anchor.Col1,
                                //图片所在结束列
                                MaxCol = anchor.Col2,
                                //图片数据
                                FileBytes = picture.PictureData.Data,
                                //文件名称
                                FileName = picture.FileName,
                                //响应类型
                                MimeType = picture.PictureData.MimeType,
                                //扩展名称
                                ExtensionName = picture.PictureData.SuggestFileExtension(),
                                //图片索引
                                FileIndex = picture.PictureIndex
                            };
                            fileInfoList.Add(entity);
                        }
                    }
                }
            }

            return fileInfoList;
        }

        /// <summary>
        /// 获取所有的图片信息
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="workSpace"></param>
        /// <returns></returns>
        private static List<ColumnFile> GetAllPictureInfos(XSSFSheet sheet,WorkSpace workSpace)
        {
            List<ColumnFile> fileInfoList = new List<ColumnFile>();

            var documentPartList = sheet.GetRelations();
            foreach (var documentPart in documentPartList)
            {
                if (documentPart is XSSFDrawing)
                {
                    var drawing = (XSSFDrawing)documentPart;
                    var shapeList = drawing.GetShapes();
                    foreach (var shape in shapeList)
                    {
                        if (shape is XSSFPicture)
                        {
                            var picture = (XSSFPicture)shape;
                            var anchor = picture.GetAnchor();
                            //var anchor = picture.GetPreferredSize();

                            int row1 = (int)anchor.GetType().GetProperty("Row1").GetValue(anchor);
                            int row2;
                            int col1 = (int)anchor.GetType().GetProperty("Col1").GetValue(anchor);
                            int col2;
                            try
                            {
                                row2 = (int)anchor.GetType().GetProperty("Row2").GetValue(anchor);
                                col2 = (int)anchor.GetType().GetProperty("Col2").GetValue(anchor);
                            }
                            catch
                            {
                                row2 = row1;//给默认值
                                col2 = col1;//给默认值       
                            }

                            if (IsInternalOrIntersect(row1, row2, col1, col2, workSpace))
                            {
                                ColumnFile entity = new ColumnFile()
                                {
                                    //图片所在开始行
                                    MinRow = row1,
                                    //图片所在结束行
                                    MaxRow = row2,
                                    // 图片所在开始列
                                    MinCol = col1,
                                    //图片所在结束列
                                    MaxCol = col2,
                                    //图片数据
                                    FileBytes = picture.PictureData.Data,
                                    //文件名称
                                    //FileName = picture.FileName,
                                    //响应类型
                                    MimeType = picture.PictureData.MimeType,
                                    //扩展名称
                                    ExtensionName = picture.PictureData.SuggestFileExtension(),
                                    //图片索引
                                    //FileIndex = picture.PictureIndex
                                };
                                fileInfoList.Add(entity);
                            }
                        }
                    }
                }
            }

            return fileInfoList;
        }

        /// <summary>
        /// 当前图片的位置
        /// </summary>
        /// <param name="pictureMinRow"></param>
        /// <param name="pictureMaxRow"></param>
        /// <param name="pictureMinCol"></param>
        /// <param name="pictureMaxCol"></param>
        /// <param name="workSpace"></param>
        /// <returns></returns>
        private static bool IsInternalOrIntersect(
            int pictureMinRow, int pictureMaxRow, int pictureMinCol, int pictureMaxCol, WorkSpace workSpace)
        {
            int _rangeMinRow = workSpace.MinRow ?? pictureMinRow;
            int _rangeMaxRow = workSpace.MaxRow ?? pictureMaxRow;
            int _rangeMinCol = workSpace.MinRow ?? pictureMinCol;
            int _rangeMaxCol = workSpace.MaxCol ?? pictureMaxCol;

            if (workSpace.OnlyInternal)
            {
                return (_rangeMinRow <= pictureMinRow && _rangeMaxRow >= pictureMaxRow &&
                        _rangeMinCol <= pictureMinCol && _rangeMaxCol >= pictureMaxCol);
            }
            else
            {
                return ((Math.Abs(_rangeMaxRow - _rangeMinRow) + Math.Abs(pictureMaxRow - pictureMinRow) >= Math.Abs(_rangeMaxRow + _rangeMinRow - pictureMaxRow - pictureMinRow)) &&
                (Math.Abs(_rangeMaxCol - _rangeMinCol) + Math.Abs(pictureMaxCol - pictureMinCol) >= Math.Abs(_rangeMaxCol + _rangeMinCol - pictureMaxCol - pictureMinCol)));
            }
        }
    }
}

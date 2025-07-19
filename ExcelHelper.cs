using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace CarrotMRO
{
    public class ExcelHelper
    {
        /// <summary>
        /// 从Excel读取数据到List<GeneralItem>
        /// </summary>
        /// <param name="filePath">Excel文件路径</param>
        /// <returns>读取的项目列表</returns>
        public static List<GeneralItem> ReadFromExcel(string filePath, Func<IXLWorksheet, int, GeneralItem> factory)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Excel文件不存在: {filePath}");
            }

            var items = new List<GeneralItem>();

            try
            {
                using var workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheet(1); // 获取第一个工作表
                var range = worksheet.RangeUsed(); // 获取使用的范围

                // 跳过表头行，从第二行开始读取
                for (int row = 2; !string.IsNullOrEmpty(worksheet.Cell(row, 1).GetString()); row++)
                {
                    //var item = new GeneralItem
                    //{
                    //    Part = worksheet.Cell(row, 1).GetString(),
                    //    CustomName = worksheet.Cell(row, 2).GetString(),
                    //    Name = worksheet.Cell(row, 2).GetString(),
                    //    Description = worksheet.Cell(row, 3).GetString(),
                    //    PerPrice = worksheet.Cell(row, 4).GetValue<double>()
                    //};
                    var item = factory(worksheet, row);

                    items.Add(item);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"读取Excel文件时出错: {ex.Message}", ex);
            }

            return items;
        }

        /// <summary>
        /// 将List<GeneralItem>写入Excel
        /// </summary>
        /// <param name="items">要写入的项目列表</param>
        /// <param name="filePath">Excel文件路径</param>
        /// <param name="overwrite">是否自动覆盖已存在文件</param>
        /// <returns>是否成功写入</returns>
        public static bool WriteToExcel(List<GeneralItem> items, string filePath, bool overwrite, string[] header, Action<IXLWorksheet, int, GeneralItem> itemFactory)
        {
            if (items == null)
            {
                throw new ArgumentException();
            }

            // 检查文件是否存在且不允许覆盖
            if (File.Exists(filePath) && !overwrite)
            {
                return false;
            }

            try
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Items");

                // 写入表头
                for (int i = 0; i < header.Length; i++)
                    worksheet.Cell(1, i + 1).Value = header[i];

                // 写入数据
                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    var row = i + 2; // 数据从第二行开始

                    //worksheet.Cell(row, 1).Value = item.ID;
                    //worksheet.Cell(row, 2).Value = item.Name;
                    //worksheet.Cell(row, 3).Value = item.Description;
                    //worksheet.Cell(row, 4).Value = item.Price;
                    //worksheet.Cell(row, 5).Value = item.Quantity;
                    //worksheet.Cell(row, 6).Value = item.CreatedDate;

                    itemFactory(worksheet, row, item);
                }

                // 自动调整列宽
                //worksheet.Columns().AdjustToContents();

                // 保存文件
                workbook.SaveAs(filePath);
                return true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"写入Excel文件时出错: {ex.Message}", ex);
            }
        }
        public static bool WriteToExcel(IEnumerable<IGrouping<string, GeneralItem>> groupedItems, string filePath, bool overwrite, string[] header, Action<IXLWorksheet, int, GeneralItem> itemFactory)
        {
            if (groupedItems == null)
            {
                throw new ArgumentException();
            }

            // 检查文件是否存在且不允许覆盖
            if (File.Exists(filePath) && !overwrite)
            {
                return false;
            }

            try
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Report");

                // 写入表头
                for (int i = 0; i < header.Length; i++)
                    worksheet.Cell(1, i + 1).Value = header[i];

                // 写入数据
                int row = 2; // 数据从第二行开始
                foreach (var group in groupedItems)
                {
                    worksheet.Cell(row, 1).Value = group.Key;

                    // 合并
                    var rangeToMerge = worksheet.Range(row, 1, row, header.Length);
                    rangeToMerge.Merge();
                    rangeToMerge.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rangeToMerge.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    
                    row++;

                    foreach (var item in group)
                    {
                        itemFactory(worksheet, row, item);
                        row++;
                    }
                }

                // 自动调整列宽
                //worksheet.Columns().AdjustToContents();

                // 保存文件
                workbook.SaveAs(filePath);
                return true;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"写入Excel文件时出错: {ex.Message}", ex);
            }
        }
    }
}
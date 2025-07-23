using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CarrotMRO
{
    public static class ExcelItemFactory
    {
        // TODO
        public static string[] AutoSaveHeader = new string[] {
            "类目",
            "自定项目",
            "标准项目",
            "单位",
            "数量",
            "单价",
            "备注" };

        //public static string[] OutHeader = new string[] {
        //    "标准项目",
        //    "单位",
        //    "数量",
        //    "单价",
        //    "总价",
        //    "备注" };

        public static GeneralItem StandardItemRead(AppConfig config, IXLWorksheet worksheet, int row)
        {
            //var item = new GeneralItem {
            //    Name = worksheet.Cell(row, 1).GetString(),
            //    Unit = worksheet.Cell(row, 2).GetString(),
            //    PerPrice = worksheet.Cell(row, 4).GetValue<double>(),
            //    Description = worksheet.Cell(row, 5).GetString()
            //};
            //return item;
            var header = config.Standard?.Header;
            if (header == null) return null;

            return new GeneralItem {
                Part = GetCellValue<string>(worksheet, row, header.PartColumn),
                CustomName = GetCellValue<string>(worksheet, row, header.CustomNameColumn),
                Name = GetCellValue<string>(worksheet, row, header.NameColumn),
                Unit = GetCellValue<string>(worksheet, row, header.UnitColumn),
                Num = GetCellValue<double>(worksheet, row, header.NumColumn),
                PerPrice = GetCellValue<double>(worksheet, row, header.PerPriceColumn),
                //SumPrice = GetCellValue<double>(worksheet, row, header.SumPriceColumn),
                Description = GetCellValue<string>(worksheet, row, header.DescColumn)
            };
        }

        public static GeneralItem AutosaveItemRead(AppConfig config, IXLWorksheet worksheet, int row)
        {
            // TODO
            var item = new GeneralItem {
                Part = worksheet.Cell(row, 1).GetString(),
                CustomName = worksheet.Cell(row, 2).GetString(),
                Name = worksheet.Cell(row, 3).GetString(),
                Unit = worksheet.Cell(row, 4).GetString(),
                Num = worksheet.Cell(row, 5).GetValue<double>(),
                PerPrice = worksheet.Cell(row, 6).GetValue<double>(),
                Description = worksheet.Cell(row, 7).GetString()
            };
            return item;
        }

        public static void AutosaveItemWrite(AppConfig config, IXLWorksheet worksheet, int row, GeneralItem item)
        {
            // TODO
            worksheet.Cell(row, 1).Value = item.Part;
            worksheet.Cell(row, 2).Value = item.CustomName;
            worksheet.Cell(row, 3).Value = item.Name;
            worksheet.Cell(row, 4).Value = item.Unit;
            worksheet.Cell(row, 5).Value = item.Num;
            worksheet.Cell(row, 6).Value = item.PerPrice;
            worksheet.Cell(row, 7).Value = item.Description;
        }

        public static void QuotationItemWrite(AppConfig config, IXLWorksheet worksheet, int row, GeneralItem item)
        {
            //worksheet.Cell(row, 1).Value = item.Name;
            //worksheet.Cell(row, 2).Value = item.Unit;
            //worksheet.Cell(row, 3).Value = item.Num;
            //worksheet.Cell(row, 4).Value = item.PerPrice;
            //worksheet.Cell(row, 5).FormulaA1 = $"=C{row}*D{row}"; // E 列（公式计算）
            //worksheet.Cell(row, 6).Value = item.Description;

            // 获取配置中的列映射
            var header = config.Quotation?.Header;
            if (header == null) return;

            // 动态设置单元格值
            SetCellValue(worksheet, row, header.NameColumn, item.Name);
            SetCellValue(worksheet, row, header.UnitColumn, item.Unit);
            SetCellValue(worksheet, row, header.NumColumn, item.Num);
            SetCellValue(worksheet, row, header.PerPriceColumn, item.PerPrice);
            if (header.NumColumn.HasValue && header.PerPriceColumn.HasValue)
            {
                SetCellFormula(worksheet, row, header.SumPriceColumn, $"={GetColumnLetter(header.NumColumn)}{row}*{GetColumnLetter(header.PerPriceColumn)}{row}");
            }
            else
            {
                SetCellValue(worksheet, row, header.SumPriceColumn, item.SumPrice);
            }
            SetCellValue(worksheet, row, header.DescColumn, item.Description);

        }

        private static T GetCellValue<T>(IXLWorksheet ws, int row, int? column)
        {
            if (!column.HasValue || column <= 0) return default;

            try
            {
                var cell = ws.Cell(row, column.Value);
                return cell.GetValue<T>();
            }
            catch
            {
                return default;
            }
        }
        private static void SetCellValue(IXLWorksheet ws, int row, int? col, XLCellValue value)
        {
            if (!col.HasValue || col <= 0) return;
            ws.Cell(row, col.Value).Value = value;
        }

        private static void SetCellFormula(IXLWorksheet ws, int row, int? col, string formula)
        {
            if (!col.HasValue || col <= 0 || string.IsNullOrEmpty(formula)) return;
            ws.Cell(row, col.Value).FormulaA1 = formula;
        }

        private static string GetColumnLetter(int? columnNumber)
        {
            if (!columnNumber.HasValue) return "A";
            return ((char)('A' + columnNumber.Value - 1)).ToString();
        }
    }
}

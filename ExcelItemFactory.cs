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
        public static GeneralItem ItemRead(HeaderConfig headers, IXLWorksheet worksheet, int row)
        {
            // Standard Old Generation
            //var item = new GeneralItem {
            //    Name = worksheet.Cell(row, 1).GetString(),
            //    Unit = worksheet.Cell(row, 2).GetString(),
            //    PerPrice = worksheet.Cell(row, 4).GetValue<double>(),
            //    Description = worksheet.Cell(row, 5).GetString()
            //};
            //return item;
            // Autosave Old Generation
            //var item = new GeneralItem
            //{
            //    Part = worksheet.Cell(row, 1).GetString(),
            //    CustomName = worksheet.Cell(row, 2).GetString(),
            //    Name = worksheet.Cell(row, 3).GetString(),
            //    Unit = worksheet.Cell(row, 4).GetString(),
            //    Num = worksheet.Cell(row, 5).GetValue<double>(),
            //    PerPrice = worksheet.Cell(row, 6).GetValue<double>(),
            //    Description = worksheet.Cell(row, 7).GetString()
            //};
            //return item;

            if (headers == null) return null;

            return new GeneralItem
            {
                Part = GetCellValue<string>(worksheet, row, headers.PartColumn),
                CustomName = GetCellValue<string>(worksheet, row, headers.CustomNameColumn),
                Name = GetCellValue<string>(worksheet, row, headers.NameColumn),
                Unit = GetCellValue<string>(worksheet, row, headers.UnitColumn),
                Num = GetCellValue<double>(worksheet, row, headers.NumColumn),
                PerPrice = GetCellValue<double>(worksheet, row, headers.PerPriceColumn),
                MaterialInPerPrice = GetCellValue<double>(worksheet, row, headers.MaterialInPerPriceColumn),
                BaseMaterialInPerPrice = GetCellValue<double>(worksheet, row, headers.BaseMaterialInPerPriceColumn),
                AuxMaterialInPerPrice = GetCellValue<double>(worksheet, row, headers.AuxMaterialInPerPriceColumn),
                MachineInPerPrice = GetCellValue<double>(worksheet, row, headers.MachineInPerPriceColumn),
                LaborInPerPrice = GetCellValue<double>(worksheet, row, headers.LaborInPerPriceColumn),
                SumPrice = GetCellValue<double>(worksheet, row, headers.SumPriceColumn),
                Description = GetCellValue<string>(worksheet, row, headers.DescColumn)
            };
        }

        public static void ItemWrite(HeaderConfig headers, IXLWorksheet worksheet, int row, GeneralItem item)
        {
            // Autosave Old Generation
            //worksheet.Cell(row, 1).Value = item.Part;
            //worksheet.Cell(row, 2).Value = item.CustomName;
            //worksheet.Cell(row, 3).Value = item.Name;
            //worksheet.Cell(row, 4).Value = item.Unit;
            //worksheet.Cell(row, 5).Value = item.Num;
            //worksheet.Cell(row, 6).Value = item.PerPrice;
            //worksheet.Cell(row, 7).Value = item.Description;

            // Quotation Old Generation
            //worksheet.Cell(row, 1).Value = item.Name;
            //worksheet.Cell(row, 2).Value = item.Unit;
            //worksheet.Cell(row, 3).Value = item.Num;
            //worksheet.Cell(row, 4).Value = item.PerPrice;
            //worksheet.Cell(row, 5).FormulaA1 = $"=C{row}*D{row}"; // E 列（公式计算）
            //worksheet.Cell(row, 6).Value = item.Description;

            // 获取配置中的列映射
            if (headers == null) return;

            // 动态设置单元格值
            SetCellValue(worksheet, row, headers.NameColumn, item.Name);
            SetCellValue(worksheet, row, headers.UnitColumn, item.Unit);
            SetCellValue(worksheet, row, headers.NumColumn, item.Num);
            SetCellValue(worksheet, row, headers.BaseMaterialInPerPriceColumn, item.BaseMaterialInPerPrice);
            SetCellValue(worksheet, row, headers.AuxMaterialInPerPriceColumn, item.AuxMaterialInPerPrice);
            SetCellValue(worksheet, row, headers.MachineInPerPriceColumn, item.MachineInPerPrice);
            SetCellValue(worksheet, row, headers.LaborInPerPriceColumn, item.LaborInPerPrice);

            if (headers.BaseMaterialInPerPriceColumn.HasValue
                && headers.AuxMaterialInPerPriceColumn.HasValue)
            {
                // 如果有主材与辅材列,则设置公式计算材料单价
                SetCellFormula(worksheet, row, headers.MaterialInPerPriceColumn,
                    $"={GetColumnLetter(headers.BaseMaterialInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.AuxMaterialInPerPriceColumn)}{row}");
            }
            else
            {
                SetCellValue(worksheet, row, headers.MaterialInPerPriceColumn, item.MaterialInPerPrice);
            }

            if (headers.MaterialInPerPriceColumn.HasValue
                && headers.MachineInPerPriceColumn.HasValue
                && headers.LaborInPerPriceColumn.HasValue)
            {
                // 如果有材料、机械、人工列,则设置公式计算综合单价
                SetCellFormula(worksheet, row, headers.PerPriceColumn,
                    $"={GetColumnLetter(headers.MaterialInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.MachineInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.LaborInPerPriceColumn)}{row}");
            }
            else if (headers.BaseMaterialInPerPriceColumn.HasValue
                && headers.AuxMaterialInPerPriceColumn.HasValue
                && headers.MachineInPerPriceColumn.HasValue
                && headers.LaborInPerPriceColumn.HasValue)
            {
                // 如果有主材、辅材、机械、人工列,则设置公式计算综合单价
                SetCellFormula(worksheet, row, headers.PerPriceColumn,
                    $"={GetColumnLetter(headers.BaseMaterialInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.AuxMaterialInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.MachineInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.LaborInPerPriceColumn)}{row}");
            }
            else
            {
                SetCellValue(worksheet, row, headers.PerPriceColumn, item.PerPrice);
            }

            if (headers.NumColumn.HasValue
                && headers.PerPriceColumn.HasValue)
            {
                // 如果有数量与单价列,则设置公式计算总价
                SetCellFormula(worksheet, row, headers.SumPriceColumn,
                    $"={GetColumnLetter(headers.NumColumn)}{row}" +
                    $"*{GetColumnLetter(headers.PerPriceColumn)}{row}");
            }
            else if (headers.NumColumn.HasValue
                && headers.MaterialInPerPriceColumn.HasValue
                && headers.MachineInPerPriceColumn.HasValue
                && headers.LaborInPerPriceColumn.HasValue)
            {
                // 如果有数量与材料、机械、人工列,则设置公式计算总价
                SetCellFormula(worksheet, row, headers.SumPriceColumn,
                    $"={GetColumnLetter(headers.NumColumn)}{row}" +
                    $"*(" +
                    $"{GetColumnLetter(headers.MaterialInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.MachineInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.LaborInPerPriceColumn)}{row}" +
                    ")");
            }
            else if (headers.NumColumn.HasValue
                && headers.BaseMaterialInPerPriceColumn.HasValue
                && headers.AuxMaterialInPerPriceColumn.HasValue
                && headers.MachineInPerPriceColumn.HasValue
                && headers.LaborInPerPriceColumn.HasValue)
            {
                // 如果有数量与主材、辅材、机械、人工列,则设置公式计算总价
                SetCellFormula(worksheet, row, headers.SumPriceColumn,
                    $"={GetColumnLetter(headers.NumColumn)}{row}" +
                    $"*(" +
                    $"{GetColumnLetter(headers.BaseMaterialInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.AuxMaterialInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.MachineInPerPriceColumn)}{row}" +
                    $"+{GetColumnLetter(headers.LaborInPerPriceColumn)}{row}" +
                    ")");
            }
            else
            {
                SetCellValue(worksheet, row, headers.SumPriceColumn, item.SumPrice);
            }
            SetCellValue(worksheet, row, headers.DescColumn, item.Description);

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

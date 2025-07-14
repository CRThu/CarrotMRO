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
        public static string[] AutoSaveHeader = new string[] {
            "类目",
            "自定项目",
            "标准项目",
            "单位",
            "数量",
            "单价",
            "备注" };

        public static GeneralItem StandardReadFactory(IXLWorksheet worksheet, int row)
        {
            var item = new GeneralItem
            {
                Name = worksheet.Cell(row, 1).GetString(),
                Unit = worksheet.Cell(row, 2).GetString(),
                PerPrice = worksheet.Cell(row, 4).GetValue<double>()
            };
            return item;
        }

        public static GeneralItem AutosaveReadFactory(IXLWorksheet worksheet, int row)
        {
            var item = new GeneralItem
            {
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

        public static void AutosaveWriteFactory(IXLWorksheet worksheet, int row, GeneralItem item)
        {

            worksheet.Cell(row, 1).Value = item.Part;
            worksheet.Cell(row, 2).Value = item.CustomName;
            worksheet.Cell(row, 3).Value = item.Name;
            worksheet.Cell(row, 4).Value = item.Unit;
            worksheet.Cell(row, 5).Value = item.Num;
            worksheet.Cell(row, 6).Value = item.PerPrice;
            worksheet.Cell(row, 7).Value = item.Description;
        }
    }
}

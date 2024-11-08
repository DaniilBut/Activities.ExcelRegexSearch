using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using BR.Core;
using BR.Core.Attributes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Activities.ExcelRegexSearch
{
    [ScreenName("Поиск по регулярному выражению в Excel")]
    [BR.Core.Attributes.Path("Custom activities")]
    public class ExcelRegexSearch : Activity
    {
        [ScreenName("Путь к файлу")]
        [Description("Путь к Excel файлу")]
        [IsRequired]
        public string str_filePath { get; set; }

        [ScreenName("Имя листа")]
        [Description("Имя листа в Excel")]
        [IsRequired]
        public string str_sheetName { get; set; }

        [ScreenName("Регулярное выражение")]
        [Description("Шаблон для поиска")]
        [IsRequired]
        public string str_regexPattern { get; set; }

        [ScreenName("Диапазон поиска")]
        [Description("Диапазон ячеек для поиска")]
        public string str_searchRange { get; set; }

        [ScreenName("Результаты")]
        [Description("Список найденных адресов ячеек")]
        [IsOut]
        public List<string> str_cellAddresses { get; set; }

        public override void Execute(int? optionID)
        {
            str_cellAddresses = new List<string>();
            var excelApp = new Excel.Application();
            var workbook = excelApp.Workbooks.Open(str_filePath);
            var worksheet = (Excel.Worksheet)workbook.Worksheets[str_sheetName];
            Excel.Range range;
            if (string.IsNullOrEmpty(str_searchRange))
            {
                range = worksheet.UsedRange;
            }
            else if (str_searchRange.Contains(":"))
            {
                range = worksheet.get_Range(str_searchRange);
            }
            else if (int.TryParse(str_searchRange, out int startRow))
            {
                int lastRow = worksheet.UsedRange.Rows.Count;
                string address = $"A{startRow}:Z{lastRow}";
                range = worksheet.get_Range(address);
            }
            else
            {
                throw new ArgumentException("Неверный формат диапазона поиска.");
            }

            Regex regex = new Regex(str_regexPattern);

            foreach (Excel.Range cell in range)
            {
                if (cell.Value2 != null && regex.IsMatch(cell.Value2.ToString()))
                {
                    str_cellAddresses.Add(cell.Address[false, false]);
                }
            }

            workbook.Close(false);
            Marshal.ReleaseComObject(workbook);
            excelApp.Quit();
            Marshal.ReleaseComObject(excelApp);
        }
    }
}
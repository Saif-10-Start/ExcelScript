using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExcelScript;
internal class Program
{
    static void Main(string[] args)
    {
        var startingTime = DateTime.Now;
        using var wb = new XLWorkbook("C:\\Users\\Ready2Go\\OneDrive - Detmold\\Dokumente\\Work\\ExcelScript\\test.xlsx");
        using var newWb = new XLWorkbook();

        foreach (var sheet in wb.Worksheets)
        {
            int col = 0;
            int row = 0;

            if (sheet.Name == "Summary List") continue;
            newWb.AddWorksheet(sheet.Name);

            foreach (var column in sheet.ColumnsUsed(options: XLCellsUsedOptions.Contents))
            {
                if (column.IsMainColumn())
                {
                    col += 1;
                    row = 0;
                    var header = Extensions.GetHeader(column.FirstCellUsed().Value.ToString());

                    foreach (var cell in column.CellsUsed())
                    {
                        row += 1;

                        var newCell = newWb.Worksheet(sheet.Name).Cell(row, col); ;
                        newCell.SetValue(cell.Value);

                        if (!Extensions.GetHeaderNames().Contains(newCell.Value.ToString()) && newCell.Address.RowNumber != 1)
                        {
                            // ToDo: NOT WORKING AT ALLLLLLLL...
                            var value = cell.Value.ToString().Trim();
                            cell.Value = Regex.Replace(value, @"(?i)^n[./\-]?a$", "");

                            switch (header)
                            {
                                case HeaderName.Website:
                                    newCell.CleanLink();
                                    break;

                                case HeaderName.PhoneNumber:
                                    newCell.FormatPhoneNumber();
                                    break;

                                case HeaderName.Email:
                                    newCell.FormatMail();
                                    break;

                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        foreach (var sheet in newWb.Worksheets)
        {
            int width = sheet.ColumnsUsed().Count();
            int height = sheet.RowsUsed().Count();

            for (int i = 1; i <= width; i++)
            {
                for (int j = 1; j <= height; j++)
                {
                    var cell = sheet.Cell(j, i);
                    if (j == 1)
                    {
                        cell.Style.Fill.SetBackgroundColor(XLColor.AirForceBlue);
                        cell.Style.Font.FontSize = 14;
                        cell.Style.Font.SetBold();
                        cell.Value = cell.Value.ToString().FirstCharToUpper();
                    }
                    else if (cell.Address.RowNumber % 2 == 0) cell.Style.Fill.SetBackgroundColor(XLColor.AliceBlue);
                    else cell.Style.Fill.SetBackgroundColor(XLColor.White);
                }
            }

            sheet.ColumnsUsed().AdjustToContents();
        }

        Console.WriteLine("Saving the new Table...");
        newWb.SaveAs("temp.xlsx");
        Console.WriteLine("Done!");
        Console.WriteLine($"Everything finished in: {DateTime.Now - startingTime}");
    }
}
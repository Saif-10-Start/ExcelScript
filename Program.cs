using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelScript;
internal class Program
{
    static void Main(string[] args)
    {
        var startingTime = DateTime.Now;

        XLColor HeaderColor = XLColor.AirForceBlue;
        XLColor Alternating1 = XLColor.AliceBlue;
        XLColor Alternating2 = XLColor.White;

        using var wb = new XLWorkbook("..\\..\\..\\test.xlsx");
        using var newWb = new XLWorkbook();
        Dictionary<IXLWorksheet, IXLCell> Comments = [];

        // Copy Data
        foreach (var sheet in wb.Worksheets)
        {
            int col = 0;
            int row = 0;

            if (sheet.Name == "Summary List") continue;
            newWb.AddWorksheet(sheet.Name);

            foreach (var column in sheet.ColumnsUsed(options: XLCellsUsedOptions.Contents))
            {
                col += 1;
                row = 0;
                var header = Extensions.GetHeader(column.FirstCellUsed().Value.ToString());

                foreach (var cell in column.CellsUsed())
                {
                    row += 1;

                    if (cell.IsComment())
                    {
                        Comments.Add(sheet, cell);
                        continue;
                    }

                    var newCell = newWb.Worksheet(sheet.Name).Cell(row, col); ;
                    newCell.SetValue(cell.Value);

                    if (newCell.Address.RowNumber != 1)
                    {
                        newCell.Strip();    // Strip the cell value from any unwanted characters

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

                            case HeaderName.Location:
                                newCell.FormatLocation();
                                break;

                            default:
                                break;
                        }
                    }
                }
            }
        }

        // Styling the new Table
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
                        cell.Style.Fill.SetBackgroundColor(HeaderColor);
                        cell.Style.Font.FontSize = 14;
                        cell.Style.Font.SetBold();
                        cell.Value = cell.Value.ToString().FirstCharToUpper();
                    }
                    else if (cell.Address.RowNumber % 2 == 0) cell.Style.Fill.SetBackgroundColor(Alternating1);
                    else cell.Style.Fill.SetBackgroundColor(Alternating2);
                }
            }

            sheet.ColumnsUsed().AdjustToContents();
        }

        // Copy over the comments with their styling
        foreach (var comment in Comments)
        {
            var newCell = newWb.Worksheet(comment.Key.Name).Cell(comment.Value.Address.RowNumber, comment.Value.Address.ColumnNumber);
            newCell.SetValue(comment.Value.Value);
            newCell.Style = comment.Value.Style;
        }

        // Add Sumary List Sheet
        var summary = newWb.AddWorksheet("Summary List");
        string[] categories = ["Category", "Number of Contractors", "Number of Physical shops", "Providing Physical Installation", "Importing Supply", "Local Supply", "Provide Delivery"];
        var sheets = newWb.Worksheets.Where(s => s.Name != "Summary List").ToArray();

        foreach (var category in categories)
        {
            var cell = summary.Cell(1, Array.IndexOf(categories, category) + 1);
            cell.SetValue(category);
            cell.Style.Fill.SetBackgroundColor(HeaderColor);
        }
        foreach (var sheet in sheets)
        {
            var cell = summary.Cell(Array.IndexOf(sheets, sheet) + 2, 1);
            cell.SetValue(sheet.Name);
        }


        Console.WriteLine("Saving the new Table...");
        newWb.SaveAs("..\\..\\..\\OutPut.xlsx");
        Console.WriteLine("Done!");
        Console.WriteLine($"Everything finished in: {DateTime.Now - startingTime}");
    }
}
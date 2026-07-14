using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

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

        // Add Sumary List Sheet
        var summary = newWb.AddWorksheet("Summary List", 1);
        string[] categories = ["Category", "Number of Contractors", "Number of Physical stores", "Providing Physical Installation", "Importing Supply", "Local Supply", "Provide Delivery"];
        var sheets = newWb.Worksheets.Where(s => s.Name != "Summary List").ToArray();

        foreach (var category in categories)
        {
            var cell = summary.Cell(1, Array.IndexOf(categories, category) + 1);
            cell.SetValue(category);
        }
        foreach (var sheet in sheets)
        {
            var row = Array.IndexOf(sheets, sheet) + 2;

            // Set BackLink to the summary
            var col = sheet.Row(1).CellsUsed().Last().Address.ColumnNumber + 1;
            var cell = sheet.Cell(1, col);
            cell.SetFormulaA1($"=HYPERLINK(\"#'Summary List'!A{row}\", \"Back to Summary\")");
            cell.Style.Font.SetUnderline().Font.FontColor = XLColor.Blue;

            // Set Category
            SetFormula(row, 1, $"=HYPERLINK(\"#'{sheet.Name}'!A1\", \"{sheet.Name}\")");
            summary.Cell(row, 1).Style.Font.SetUnderline().Font.FontColor = XLColor.Blue;

            // Set Number of Contractors
            SetFormula(row, 2, $"=COUNTA('{sheet.Name}'!A:A)-1");

            // Set Number of Physical Shops
            SetCounter(row, 3, "Physical Store", "B" + row, sheet.Name);

            // Set Providing Physical Installation
            SetCounter(row, 4, "Installation", "B" + row, sheet.Name);

            // Set Importing Supply
            SetCounter(row, 5, "Import", "B" + row, sheet.Name);

            // Set Local Supply
            SetCounter(row, 6, "Local Supply", "B" + row, sheet.Name);

            // Set Provide Delivery
            SetCounter(row, 7, "Delivery", "B" + row, sheet.Name);
        }
        foreach (var column in summary.ColumnsUsed())
            column.AdjustToContents();

        void SetFormula(int row, int column, string formula)
        {
            summary.Cell(row, column).FormulaA1 = formula;
            summary.Cell(row, column).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        void SetCounter(int row, int column, string match, string numberOfField, string sheetName)
            => SetFormula(row, column, $"=IFERROR(COUNTIF(INDEX('{sheetName}'!$A:$Z, 0, MATCH(\"{match}\", '{sheetName}'!$1:$1, 0)), \"yes\") & \" out of \" & {numberOfField}, \"—\")");


        // Styling the new Table
        foreach (var sheet in newWb.Worksheets)
        {
            var rows = sheet.Rows();
            var height = sheet.Rows().Count() + 50;
            for (int i = 1; i <= height; i++)
            {
                var row = sheet.Row(i);
                if (i == 1)
                {
                    row.Style.Fill.SetBackgroundColor(HeaderColor);
                    row.Style.Font.FontSize = 14;
                    row.Style.Font.SetBold();

                    foreach (var cell in row.CellsUsed())
                        if (!cell.HasFormula) cell.Value = cell.Value.ToString().FirstCharToUpper();
                }
                else if (row.RowNumber() % 2 == 0) row.Style.Fill.SetBackgroundColor(Alternating1);
                else row.Style.Fill.SetBackgroundColor(Alternating2);
            }

            sheet.Columns().AdjustToContents();
        }

        // Copy over the comments with their styling
        foreach (var comment in Comments)
        {
            var newCell = newWb.Worksheet(comment.Key.Name).Cell(comment.Value.Address.RowNumber, comment.Value.Address.ColumnNumber);
            newCell.SetValue(comment.Value.Value);
            newCell.Style = comment.Value.Style;
        }

        Console.WriteLine("Saving the new Table...");
        newWb.SaveAs("..\\..\\..\\Output.xlsx");
        Console.WriteLine("Done!");
        Console.WriteLine($"Everything finished in: {DateTime.Now - startingTime}");
    }
}
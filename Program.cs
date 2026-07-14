using ClosedXML.Excel;

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

        List<(string SheetName, int RowNumber)> rowsToInsert = new();

        // 1. Copy Raw Data (Without formatting or native hyperlinks to avoid shifting bugs)
        foreach (var sheet in wb.Worksheets)
        {
            if (sheet.Name == "Summary List") continue;
            newWb.AddWorksheet(sheet.Name);
            var targetSheet = newWb.Worksheet(sheet.Name);

            int sourceItemCol = -1;
            int searchColIdx = 1;

            foreach (var firstRowCell in sheet.FirstRowUsed().Cells())
            {
                var header = Extensions.GetHeader(firstRowCell.Value.ToString());
                if (header == HeaderName.Item)
                {
                    sourceItemCol = searchColIdx;
                    break;
                }
                searchColIdx++;
            }
            if (sourceItemCol != -1)
            {
                var srcLastRow = sheet.LastRowUsed().RowNumber();
                var srcLastCol = sheet.LastColumnUsed().ColumnNumber();
                if (srcLastRow > 1)
                {
                    var srcRange = sheet.Range(2, 1, srcLastRow, srcLastCol);
                    srcRange.Sort(sourceItemCol);
                }
            }

            int col = 0;
            int itemColIndex = -1; // Track which column column is the "Item" column

            // COPY ALL COLUMNS FIRST
            foreach (var column in sheet.ColumnsUsed(options: XLCellsUsedOptions.Contents))
            {
                col += 1;
                int row = 0;
                var header = Extensions.GetHeader(column.FirstCellUsed().Value.ToString());

                if (header == HeaderName.Item)
                {
                    itemColIndex = col;
                }

                foreach (var cell in column.CellsUsed())
                {
                    row += 1;

                    if (cell.IsComment())
                    {
                        Comments.Add(sheet, cell);
                        continue;
                    }

                    var newCell = targetSheet.Cell(row, col);
                    newCell.SetValue(cell.Value);

                    if (newCell.Address.RowNumber != 1)
                    {
                        newCell.Strip();
                    }
                }
            }

            // Detect splits on targetSheet (No sorting needed here anymore, it's already sorted!)
            if (itemColIndex != -1)
            {
                var lastRow = targetSheet.LastRowUsed()!.RowNumber();
                if (lastRow > 1)
                {
                    var itemCells = targetSheet.Column(itemColIndex).Cells(2, lastRow).ToList();
                    for (int i = 0; i < itemCells.Count - 1; i++)
                    {
                        var currentCell = itemCells[i];
                        var nextCell = itemCells[i + 1];

                        if (currentCell.Value.ToString().ToLower() != nextCell.Value.ToString().ToLower())
                        {
                            rowsToInsert.Add((targetSheet.Name, nextCell.Address.RowNumber));
                        }
                    }
                }
            }
        }

        // 2. Insert blank rows to group items together (Perfectly safe now)
        var groupedSplits = rowsToInsert
            .GroupBy(x => x.SheetName)
            .ToDictionary(g => g.Key, g => g.Select(x => x.RowNumber).OrderByDescending(r => r).ToList());

        foreach (var kvp in groupedSplits)
        {
            var sheetToModify = newWb.Worksheet(kvp.Key);
            foreach (var rowNum in kvp.Value)
            {
                sheetToModify.Row(rowNum).InsertRowsAbove(1);
            }
        }

        // 3. Post-Insert Formatting Loop (Apply CleanLink, FormatPhoneNumber, etc. after rows are finalized)
        foreach (var sheet in newWb.Worksheets)
        {
            if (sheet.Name == "Summary List") continue;

            foreach (var column in sheet.ColumnsUsed())
            {
                var header = Extensions.GetHeader(column.FirstCellUsed().Value.ToString());

                foreach (var cell in column.CellsUsed())
                {
                    if (cell.Address.RowNumber == 1) continue; // Skip header

                    switch (header)
                    {
                        case HeaderName.Website:
                            cell.CleanLink();
                            break;

                        case HeaderName.PhoneNumber:
                            cell.FormatPhoneNumber();
                            break;

                        case HeaderName.Email:
                            cell.FormatMail();
                            break;

                        case HeaderName.Location:
                            cell.FormatLocation();
                            break;

                        default:
                            break;
                    }
                }
            }
        }

        // 4. Generate the Summary List and the Hyperlinks
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


        // 5. Style the new Table
        foreach (var sheet in newWb.Worksheets)
        {
            var lastRowUsed = sheet.Cells().Count() + 50;
            for (int i = 1; i <= lastRowUsed; i++)
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

        // 6. Copy over the comments to the new workbook
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
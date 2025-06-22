using System;
using System.Drawing;
using Libs.Excel.Utils;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Libs.Excel
{
    public static class ExcelAPI
    {
        public static ExcelApp StartExcelApp(bool isVisible = false, bool displayAlerts = false)
        {
            try
            {
                var app = new ExcelApp
                {
                    DisplayAlerts = displayAlerts,
                    Visible = isVisible,
                };
                return app;
            }
            catch { }
            return null;
        }

        public static void QuitExcelApp(this ExcelApp app)
        {
            WindowUtils.GetWindowThreadProcessId(new IntPtr(app.Hwnd), out int processId);
            app.Quit();
            ProcessManager.Kill(processId);
        }

        public static Workbook OpenFile(this ExcelApp app, string filePath)
        {
            var workbooks = app.Workbooks;
            return workbooks.Open(filePath);
        }

        public static Workbook CreateNewFile(this ExcelApp app)
        {
            var workbooks = app.Workbooks;
            return workbooks.Add(System.Reflection.Missing.Value);
        }

        public static Worksheet AddSheet(this Workbook file, string sheetName, bool addLast = true)
        {
            var worksheets = file.Worksheets;
            dynamic newSheet = worksheets.Add();
            newSheet.Name = sheetName;
            if (addLast) newSheet.Move(After: file.Sheets[file.Sheets.Count]); // Chỉ số trong Excel xuất phát từ 1
            return newSheet;
        }

        /// <summary>
        /// Lấy sheet từ số thứ tự
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sheetIndex">Chỉ số trong Excel xuất phát từ 1</param>
        /// <returns></returns>
        public static Worksheet GetSheet(this Workbook file, int sheetIndex)
        {
            return file.Worksheets[sheetIndex]; // Chỉ số trong Excel xuất phát từ 1
        }

        public static void SetSheetActivate(this Workbook file, int sheetIndex)
        {
            Worksheet sheet = file.Worksheets[sheetIndex]; // Chỉ số trong Excel xuất phát từ 1
            sheet.Activate();
        }

        public static Worksheet GetLastSheet(this Workbook file)
        {
            int count = file.Worksheets.Count;
            if (count == 0) return null;
            return file.Worksheets[count]; // Chỉ số trong Excel xuất phát từ 1
        }

        public static void RenameSheet(this Worksheet sheet, string newName)
        {
            sheet.Name = newName;
        }

        public static void SetValue(this Worksheet sheet, int row, int column, object value)
        {
            Range range = sheet.Cells[row, column];
            range.Value = value;
        }

        public static object GetValue(this Worksheet sheet, int row, int column)
        {
            Range range = sheet.Cells[row, column];
            return range.Value;
        }

        public static void SetActivate(this Worksheet sheet)
        {
            sheet.Activate();
        }

        public static void Paste(this Worksheet sheet, Range range)
        {
            sheet.Paste(range);
        }

        public static Range GetRange(this Worksheet worksheet, int startColumn, int columnCount, int startRow, int rowCount)
        {
            int endColumn = startColumn + columnCount - 1;
            int endRow = startRow + rowCount - 1;
            return worksheet.Range[worksheet.Cells[startRow, startColumn], worksheet.Cells[endRow, endColumn]];
        }

        public static Range GetRows(this Range range)
        {
            return range.EntireRow; // Lấy danh sách hàng từ vùng Range
        }

        public static Range GetFullUsedRange(this Worksheet worksheet)
        {
            return worksheet.UsedRange;
        }

        public static void MergeRange(this Range range, Worksheet worksheet, XlHAlign hAlign = XlHAlign.xlHAlignCenter, XlVAlign vAlign = XlVAlign.xlVAlignCenter)
        {
            range.Merge();
            // Lấy hàng và cột của ô sau khi merge
            int mergedRow = range.Row;
            int mergedColumn = range.Column;
            Range mergedCell = worksheet.Cells[mergedRow, mergedColumn];
            mergedCell.HorizontalAlignment = hAlign;
            mergedCell.VerticalAlignment = vAlign;
        }

        public static void FreezePanesActiveWindow(this ExcelApp app, int numbOfRow)
        {
            app.ActiveWindow.SplitRow = numbOfRow;
            app.ActiveWindow.FreezePanes = true; // Cố định hàng
        }

        public static void SetFullBorder(this Range range)
        {
            range.Cells.Borders.LineStyle = XlLineStyle.xlContinuous; // Kẻ khung cho các ô trong Range
        }

        public static void SetVerticalAlignment(this Range range, XlVAlign align = XlVAlign.xlVAlignCenter)
        {
            range.VerticalAlignment = align;
        }

        public static void SetHorizontalAlignment(this Range range, XlHAlign align = XlHAlign.xlHAlignLeft)
        {
            range.HorizontalAlignment = align;
        }

        public static void SetDecimalNumber(this Range range, int decimalPlaces = 2)
        {
            string format = "0." + new string('0', decimalPlaces);
            range.NumberFormat = format;
        }

        public static void SetFontSize(this Range range, string fontName = "Arial", double? size = null)
        {
            if (!string.IsNullOrEmpty(fontName)) range.Font.Name = fontName;
            if (size.HasValue) range.Font.Size = size.Value;
        }

        public static void SetFont(this Range range, string fontName = "Arial")
        {
            if (!string.IsNullOrEmpty(fontName)) range.Font.Name = fontName;
        }

        public static void SetFontSize(this Range range, double size)
        {
            range.Font.Size = size;
        }

        public static void SetBold(this Range range)
        {
            range.Font.Bold = true;
        }

        public static void SetTextColor(this Range range, Color color)
        {
            range.Font.Color = color;
        }

        public static void SetBackgroundColor(this Range range, Color color)
        {
            range.Interior.Color = color;
        }

        public static void SetRowHeight(this Range range, double height = 20)
        {
            range.EntireRow.RowHeight = height;
        }

        public static void CopyRowHeight(this Worksheet worksheet, int fromRowIndex, int toRowIndex)
        {
            worksheet.Rows[toRowIndex].RowHeight = worksheet.Rows[fromRowIndex].RowHeight;
        }

        public static void SetColumnWidth(this Worksheet worksheet, int columnIndex, double width = 30)
        {
            worksheet.Columns[columnIndex].ColumnWidth = width;
        }

        public static void CopyColumnWidth(this Worksheet worksheet, int fromColumnIndex, int toColumnIndex)
        {
            worksheet.Columns[toColumnIndex].ColumnWidth = worksheet.Columns[fromColumnIndex].ColumnWidth;
        }

        public static void AutoFitColumnWidth(this Range range)
        {
            range.Columns.AutoFit();
        }

        public static void AddCommnent(this Range cell, string content)
        {
            cell.AddComment(content);
        }

        public static void AddHyperlinks(this Worksheet worksheet, Range range, string url, string displayText)
        {
            worksheet.Hyperlinks.Add(range, url, Type.Missing, displayText, displayText);
        }

        public static int ColumnToIndex(string column)
        {
            int index = 0;
            if (!string.IsNullOrEmpty(column))
            {
                foreach (char c in column.ToUpper())
                {
                    index = index * 26 + (c - 'A' + 1);
                }
            }
            return index;
        }
    }
}

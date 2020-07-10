using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Telerik.Windows.Controls;
using Telerik.Windows.Controls.Spreadsheet;
using Telerik.Windows.Controls.Spreadsheet.Worksheets;
using Telerik.Windows.Documents.Model;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx;
using Telerik.Windows.Documents.Spreadsheet.Formatting.FormatStrings;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Spreadsheet.Model.Protection;

namespace TelerikWpf.Test.Helpers
{
    public class SpreadSheetHelper
    {
        public static void SetStyleSheet(RadSpreadsheet grdVIEW, int SheetIndex, int intRowCnt, int intColCnt, int intFixRow = 0, int intFixCol = 0, bool isshowRowColHead = true, bool isshowGridLine = false, bool isProtected = false)
        {
            grdVIEW.Workbook.Worksheets[SheetIndex].Rows[0].SetVerticalAlignment(RadVerticalAlignment.Center); //해더 세로 중간 정렬

            // Column Autofit
            for (int i = 0; i < intColCnt; i++)
            {
                grdVIEW.Workbook.Worksheets[SheetIndex].Columns[i].SetVerticalAlignment(RadVerticalAlignment.Center); //세로 중간
                grdVIEW.Workbook.Worksheets[SheetIndex].Columns[i].AutoFitWidth(); //Autofit
                grdVIEW.Workbook.Worksheets[SheetIndex].Cells[0, i].SetHorizontalAlignment(RadHorizontalAlignment.Center);//Header만
                // grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[0, i]
            }

            // Header
            CellSelection headerSelection = grdVIEW.Workbook.Worksheets[SheetIndex].Cells[0, 0, 0, intColCnt - 1];
            ThemableColor gray = new ThemableColor(Colors.LightGray);
            ThemableColor patternColor = new ThemableColor(Colors.LightGray);
            ThemableColor backgroundColor = new ThemableColor(Colors.LightGray);
            IFill fill = new PatternFill(PatternType.Gray75Percent, patternColor, backgroundColor);
            headerSelection.SetFill(fill); //header color

            // Row
            CellSelection rowSelection = grdVIEW.Workbook.Worksheets[SheetIndex].Cells[1, 1, intRowCnt - 1, intColCnt - 1];

            // All, 
            SetBorders(grdVIEW, SheetIndex, 0, 0, intRowCnt, intColCnt - 1, Colors.Black);

            // Freeze
            if (intFixRow > 0 || intFixCol > 0)
            {
                SetFreeze(grdVIEW, SheetIndex, intFixRow, intFixCol);
            }
            // Hide or Show Row and Column Headers
            //grdVIEW.ActiveWorksheetEditor.ShowRowColumnHeadings = showRowColHead;
            grdVIEW.Workbook.Worksheets[SheetIndex].ViewState.ShowRowColHeaders = isshowRowColHead;

            // Hide or Show Gridlines
            //grdVIEW.ActiveWorksheetEditor.ShowGridlines = false;
            grdVIEW.Workbook.Worksheets[SheetIndex].ViewState.ShowGridLines = isshowGridLine;

            //Sheet Protect
            grdVIEW.Workbook.Worksheets[SheetIndex].Protect("", WorksheetProtectionOptions.Default);
        }

        // Border
        public static void SetBorders(RadSpreadsheet grdVIEW, int intCurrentIdx, int intStartRowCnt, int intStartColCnt, int intRowCnt, int intColCnt, Color color)
        {
            CellSelection selCols = grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[intStartRowCnt, intStartRowCnt, intRowCnt, intColCnt];
            ThemableColor selColor = new ThemableColor(color);
            CellBorders blackBorders = new CellBorders(
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.None, selColor),
                new CellBorder(CellBorderStyle.None, selColor)
                );
            selCols.SetBorders(blackBorders);
        }
        public static void SetBorders(RadSpreadsheet grdVIEW, int intCurrentIdx, int intStartColCnt, int intRowCnt, int intColCnt, Color color)
        {
            CellSelection selCols = grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[intRowCnt, intColCnt];
            ThemableColor selColor = new ThemableColor(color);
            CellBorders blackBorders = new CellBorders(
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.Thin, selColor),
                new CellBorder(CellBorderStyle.None, selColor),
                new CellBorder(CellBorderStyle.None, selColor)
                );
            selCols.SetBorders(blackBorders);
        }
        // Back Color
        public static void SetBackColor(RadSpreadsheet grdVIEW, int intCurrentIdx, int intRowCnt, int intColCnt, Color color)
        {
            CellSelection selCols = grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[intRowCnt, intColCnt];
            ThemableColor gray = new ThemableColor(Colors.LightGray);
            ThemableColor patternColor = new ThemableColor(Colors.LightGray);
            ThemableColor backgroundColor = new ThemableColor(Colors.LightGray);
            IFill fill = new PatternFill(PatternType.Gray75Percent, patternColor, backgroundColor);
            selCols.SetFill(fill);
        }
        public static void SetBackColor(RadSpreadsheet grdVIEW, int intCurrentIdx, int intStartRowCnt, int intStartColCnt, int intRowCnt, int intColCnt, Color color)
        {
            CellSelection selCols = grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[intStartRowCnt, intStartColCnt, intRowCnt, intColCnt];
            ThemableColor gray = new ThemableColor(color);
            ThemableColor patternColor = new ThemableColor(color);
            ThemableColor backgroundColor = new ThemableColor(color);
            IFill fill = new PatternFill(PatternType.Gray75Percent, patternColor, backgroundColor);
            selCols.SetFill(fill);
        }
        // Fore Color
        public static void SetForeColor(RadSpreadsheet grdVIEW, int intCurrentIdx, int intRowCnt, int intColCnt, Color color)
        {
            grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[intRowCnt, intColCnt].SetForeColor(new ThemableColor(color));
        }
        public static void SetForeColor(RadSpreadsheet grdVIEW, int intCurrentIdx, int intStartRowCnt, int intStartColCnt, int intRowCnt, int intColCnt, Color color)
        {
            grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[intStartRowCnt, intStartColCnt, intRowCnt, intColCnt].SetForeColor(new ThemableColor(color));
        }

        // Column Width
        public static void ExtendWidth(RadSpreadsheet grdVIEW, int sheetIndex, int colStart, int colEnd, double dblValue)
        {
            double dblCurrentColWith = 0;
            for (int i = sheetIndex; i < colEnd; i++)
            {
                if (double.TryParse(grdVIEW.Workbook.Worksheets[sheetIndex].Columns[i].GetWidth().Value.Value.ToString(), out dblCurrentColWith))
                {
                    dblCurrentColWith += dblValue;
                    grdVIEW.Workbook.Worksheets[sheetIndex].Columns[i].SetWidth(new ColumnWidth(dblCurrentColWith, true));
                }
            }
        }
        public static void SetWidth(RadSpreadsheet grdVIEW, int sheetIndex, double[] columnWidths)
        {
            int idx = 0;
            foreach (double dbl in columnWidths)
            {
                grdVIEW.Workbook.Worksheets[sheetIndex].Columns[idx].SetWidth(new ColumnWidth(dbl, true));
                idx++;
            }
        }
        public static void SetWidth(RadSpreadsheet grdVIEW, int sheetIndex, Dictionary<int, double> colWidths)
        {
            double colWith = 0;
            foreach (int key in colWidths.Keys)
            {
                if (colWidths.TryGetValue(key, out colWith))
                {
                    grdVIEW.Workbook.Worksheets[sheetIndex].Columns[key].SetWidth(new ColumnWidth(colWith, true));
                }

            }
        }
        // Height
        public static void SetHeight(RadSpreadsheet grdVIEW, int sheetIndex, int fromIndex, int toIndex, double dubHeight)
        {
            grdVIEW.Workbook.Worksheets[sheetIndex].Rows[fromIndex, toIndex].SetHeight(new RowHeight(dubHeight, true));
        }

        // Horizontal / Vertical
        public static void SetContentsHorizontalCols(RadSpreadsheet grdVIEW, int sheetIndex, int RowCount, Dictionary<int, RadHorizontalAlignment> horizonCols)
        {
            foreach (int key in horizonCols.Keys)
            {
                // grdVIEW.Workbook.Worksheets[sheetIndex].Columns[key].SetWidth(new ColumnWidth(colWith, true));

                RadHorizontalAlignment newAlign;
                if (horizonCols.TryGetValue(key, out newAlign))
                {
                    //grdVIEW.Workbook.Worksheets[sheetIndex].Columns[key].SetHorizontalAlignment(newAlign);
                    CellSelection cells = grdVIEW.Workbook.Worksheets[sheetIndex].Cells[1, key, RowCount, horizonCols.Count - 1];
                    cells.SetHorizontalAlignment(newAlign);
                }
            }
        }
        public static void SetVertical(RadSpreadsheet grdVIEW)
        {
            RadWorksheetEditor radWorksheet = (RadWorksheetEditor)grdVIEW.ActiveSheetEditor;
            if (radWorksheet != null)
                radWorksheet.VerticalScrollMode = ScrollMode.PixelBased;
        }

        // Freeze
        public static void SetFreeze(RadSpreadsheet grdVIEW, int intRowCnt, int intColCnt)
        {
            RadWorksheetEditor worksheetEditor = grdVIEW.ActiveWorksheetEditor;
            if (worksheetEditor != null)
                worksheetEditor.FreezePanes(new CellIndex(intRowCnt, intColCnt));
        }
        public static void SetFreeze(RadSpreadsheet grdVIEW, int intCurrentIdx, int intRowCnt, int intColCnt)
        {
            CellIndex fixedPaneTopLeftCellIndex = new CellIndex(intRowCnt, intColCnt);
            grdVIEW.Workbook.Worksheets[intCurrentIdx].ViewState.FreezePanes(fixedPaneTopLeftCellIndex, intRowCnt, intColCnt);
        }

        // Only display working area
        public static void SetWorkingAreaVisible(RadSpreadsheet grdVIEW, int intCurrentIdx, int intRowCnt, int intColCnt)
        {
            grdVIEW.VisibleSize = new SizeI(intRowCnt + 1, intColCnt + 1);
        }

        // Page Type (PaperTypes.A4;)
        public static void Set(RadSpreadsheet grdVIEW, PaperTypes paperTypes)
        {
            foreach (Worksheet worksheet in grdVIEW.Workbook.Worksheets)
            {
                worksheet.WorksheetPageSetup.PaperType = paperTypes;
            }
        }

        // SetWorkBook Page type


        //date
        internal static DateTime GetCellDate(CellSelection cellSelection, string strFormat)
        {
            DateTime newDate = new DateTime();
            ICellValue value = cellSelection.GetValue().Value;
            cellSelection.SetFormat(new CellValueFormat(strFormat));
            CellValueFormat format = cellSelection.GetFormat().Value;
            CellValueFormatResult formatResult = format.GetFormatResult(value);
            string result = formatResult.InfosText;
            DateTime.TryParse(result.ToString(), out newDate);
            return newDate;
        }
        // 엑셀 로드
        public Workbook LoadWorkbook(string FilePath, string type)
        {
            Workbook result = null;
            try
            {
                if (type == "01") //Resouce Excel File
                {
                    var resource = Properties.Resources._0기타업체미결;

                    if (FilePath.ToString() == "0기타업체미결.xlsx")
                    {
                        resource = Properties.Resources._0기타업체미결;
                    }
                    else
                    {
                        resource = Properties.Resources._0볼보정리;
                    }

                    XlsxFormatProvider formatProvider = new XlsxFormatProvider();
                    byte[] iStream = resource;
                    result = formatProvider.Import(iStream);
                }
                else if (type == "02") //Direct Path
                {
                    IWorkbookFormatProvider WformatProvider = new XlsxFormatProvider();
                    using (Stream input = new FileStream(FilePath, FileMode.Open))
                    {
                        result = WformatProvider.Import(input);
                    }

                }
            }
            catch (Exception ex)
            {
                return null;
            }

            return result;
        }
    }
}

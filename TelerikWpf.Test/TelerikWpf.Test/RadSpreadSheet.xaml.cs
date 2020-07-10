using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Telerik.Windows.Controls;
using Telerik.Windows.Controls.Spreadsheet;
using Telerik.Windows.Controls.Spreadsheet.Worksheets;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx;
using Telerik.Windows.Documents.Spreadsheet.Layout;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Spreadsheet.Model.Protection;
using Telerik.Windows.Documents.Spreadsheet.PropertySystem;

namespace TelerikWpf.Test
{
    public partial class RadSpreadSheet : Page
    {
        int SheetIndex = 0;
        private double verticalOffset;


        public RadSpreadSheet()
        {
            InitializeComponent();

            this.Loaded += MainWindow_Loaded;
            this.Unloaded += MainWindow_Unloaded;

            this.grdVIEW.ActiveSheetEditorChanged += this.RadSpreadsheet_ActiveSheetEditorChanged;
            this.grdVIEW.ActiveSheetEditorChanged += GrdVIEW_ActiveSheetEditorChanged;
            this.grdVIEW.Workbook.Sheets.Changed += Sheets_Changed;
            this.grdVIEW.Workbook.ActiveWorksheet.Cells.CellPropertyChanged += Cells_CellPropertyChanged;
            this.grdVIEW.WorkbookContentChanged += GrdVIEW_WorkbookContentChanged;
            SetPage();
            this.grdVIEW.MessageShowing += GrdVIEW_MessageShowing;

            btnExport.Click += BtnExport_Click;
            btnExport1.Click += BtnExport1_Click;
            btnSearch.Click += BtnSearch_Click;
            //https://www.telerik.com/forums/redirect-the-shortcut-ctrl-s
            //this.grdVIEW.ActiveWorksheetEditor.KeyBindings.RegisterCommand(myCommand, Key.S, ModifierKeys.Control);

        }

        private void GrdVIEW_MessageShowing(object sender, MessageShowingEventArgs e)
        {

        }

        private void MainWindow_Unloaded(object sender, RoutedEventArgs e)
        {
            this.grdVIEW.VerticalScrollBar.Scroll -= this.VerticalScrollBar_Scroll;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.grdVIEW.VerticalScrollBar.Scroll += this.VerticalScrollBar_Scroll;
            this.verticalOffset = this.grdVIEW.VerticalScrollBar.Value;
        }

        private void VerticalScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            Worksheet worksheet = this.grdVIEW.ActiveWorksheet;
            ViewportPane viewportPane = this.grdVIEW.ActiveWorksheetEditor.SheetViewport[ViewportPaneType.Scrollable];
            CellRange scrollablePaneVisibleRange = viewportPane.VisibleRange;
            int topRowIndex = scrollablePaneVisibleRange.FromIndex.RowIndex;
            double currentOffset = e.NewValue;
            double delta = Math.Abs(currentOffset - this.verticalOffset);

            int newRowIndex = topRowIndex;
            if (e.ScrollEventType == ScrollEventType.LargeIncrement)
            {
                double rowHeights = viewportPane.Rect.Top;
                while (rowHeights < delta)
                {
                    rowHeights += worksheet.Rows[newRowIndex].GetHeight().Value.Value;
                    newRowIndex++;
                }

                this.grdVIEW.SetVerticalOffset(rowHeights);
            }
        }
        private void GrdVIEW_WorkbookContentChanged(object sender, EventArgs e)
        {

        }

        private void Cells_CellPropertyChanged(object sender, Telerik.Windows.Documents.Spreadsheet.PropertySystem.CellPropertyChangedEventArgs e)
        {
            CellRange result = grdVIEW.ActiveWorksheet.GetUsedCellRange(new IPropertyDefinition[] { CellPropertyDefinitions.ValueProperty });
        }

        private void Sheets_Changed(object sender, SheetCollectionChangedEventArgs e)
        {

        }

        //Navigation  : The RadSpreadsheet provides a simple way to bind a key to a command. The following example shows how using the Enter key to set the cell style to "Bad"
        private void GrdVIEW_ActiveSheetEditorChanged(object sender, EventArgs e)
        {
            RadWorksheetEditor worksheetEditor = this.grdVIEW.ActiveWorksheetEditor;
            if (this.grdVIEW.ActiveWorksheetEditor != null)
            {
                worksheetEditor.KeyBindings.RegisterCommand(worksheetEditor.Commands.SetStyleCommand, Key.Enter, ModifierKeys.None, "Bad");
            }
        }

        private RadWorksheetEditor editor = null;
        private void RadSpreadsheet_ActiveSheetEditorChanged(object sender, EventArgs e)
        {
            if (this.editor != null)
            {
                this.editor.Selection.SelectionChanged -= this.Selection_SelectionChanged;
            }

            this.editor = this.grdVIEW.ActiveWorksheetEditor;

            if (this.editor != null)
            {
                this.editor.Selection.SelectionChanged += this.Selection_SelectionChanged;
            }
        }

        private void Selection_SelectionChanged(object sender, EventArgs e)
        {
            var cells = this.editor.Selection.Cells;
        }


        #region Event

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            //for (int i = 0; i < grdVIEW.Workbook.Worksheets.Count; i++)
            //{
            //    grdVIEW.Workbook.Worksheets[i].ViewState.ShowRowColHeaders = true;
            //}
            grdVIEW.ActiveWorksheetEditor.ShowRowColumnHeadings = true;

            string extension = "xlsx";
            RadSaveFileDialog SaveFileDialog = new RadSaveFileDialog()
            {
                DefaultExt = extension,
                Filter = String.Format("{1} files (*.{0})|*.{0}|All files (*.*)|*.*", extension, "Excel"),
                FilterIndex = 1
            };
            if (SaveFileDialog.ShowDialog() == true)
            {
                XlsxFormatProvider formatProvider = new XlsxFormatProvider();

                using (Stream output = new FileStream(SaveFileDialog.FileName, FileMode.Create))
                {
                    formatProvider.Export(this.grdVIEW.Workbook, output);
                }
                System.Diagnostics.Process.Start(SaveFileDialog.FileName);
            }

            //for (int i = 0; i < grdVIEW.Workbook.Worksheets.Count; i++)
            //{
            //    grdVIEW.Workbook.Worksheets[i].ViewState.ShowRowColHeaders = false;
            //}
            grdVIEW.ActiveWorksheetEditor.ShowRowColumnHeadings = false;
        }

        private void BtnExport1_Click(object sender, RoutedEventArgs e)
        {

            //않됨 위에 참조
            //RadSpreadsheet radSpreadSheet = new RadSpreadsheet();
            //radSpreadSheet.Workbook = grdVIEW.Workbook;
            //radSpreadSheet.ActiveSheet = grdVIEW.Workbook.Worksheets[grdVIEW.Workbook.ActiveWorksheet.Name];

            //ExportWhat exportSettings = new ExportWhat();
            //PrintWhatSettings settings = new PrintWhatSettings(exportSettings, false);

            //radSpreadSheet.Measure(new Size(radSpreadSheet.MinWidth, radSpreadSheet.MinHeight));
            //radSpreadSheet.Print(settings);

        }


        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            Worksheet sht = this.grdVIEW.Workbook.Worksheets[this.grdVIEW.Workbook.ActiveWorksheet.Name];
            FindOptions options = new FindOptions()
            {
                StartCell = new WorksheetCellIndex(sht, 0, 0),
                FindBy = FindBy.Rows,
                FindIn = FindInContentType.Values,
                FindWhat = "Combivent",
                FindWithin = FindWithin.Workbook,
            };

            FindResult findResult = sht.Find(options);
            IEnumerable<FindResult> findResults = sht.FindAll(options);
            //findResult.FoundCell.CellIndex
            Selection selection = this.grdVIEW.ActiveWorksheetEditor.Selection;
            selection.BeginUpdate();
            int colCnt = 0;
            foreach (FindResult result in findResults)
            {
                CellIndex cell = result.FoundCell.CellIndex;
                //CellSelection selection = sht.Cells[cell.RowIndex, cell.ColumnIndex];  
                //ThemableColor patternColor = new ThemableColor(Colors.LightGray);
                //ThemableColor backgroundColor = new ThemableColor(Colors.LightGray);
                //IFill fill = new PatternFill(PatternType.Gray75Percent, patternColor, backgroundColor);
                //selection.SetFill(fill);

                if (colCnt == 0)
                {
                    //sht.Cells[cell.ColumnIndex, cell.RowIndex].SetValue(sht.Cells[cell.ColumnIndex, cell.RowIndex].GetValue().Value.RawValue);
                    selection.Select(cell, false);
                    selection.ActiveCellMode = ActiveCellMode.Display;
                }
                else
                {
                    selection.Select(cell, false);
                    selection.ActiveCellMode = ActiveCellMode.Display;
                }
                colCnt++;
            }
            selection.EndUpdate();
        }

        #endregion

        #region User Method
        private void SetPage()
        {
            if (grdVIEW.Workbook.Worksheets.Count > 0)
            {
                foreach (Worksheet sht in grdVIEW.Workbook.Worksheets)
                {
                    //grdVIEW.Workbook.Worksheets.Remove(sht.Name);
                }
            }
            //dataset
            DataSet dataSet = InitDs();
            int coltbCount = 0;
            int rowtbCount = 0;
            for (int i = 0; i < dataSet.Tables.Count; i++)
            {
                if (i > 0)
                    grdVIEW.Workbook.Worksheets.Add();
                coltbCount = dataSet.Tables[i].Columns.Count;
                rowtbCount = dataSet.Tables[i].Rows.Count;
                SheetIndex = grdVIEW.Workbook.Worksheets.Count - 1;
                grdVIEW.Workbook.ActiveWorksheet = grdVIEW.Workbook.Worksheets[SheetIndex];
                DataTable currentTable = dataSet.Tables[i];
                Worksheet currentWorksheet = grdVIEW.Workbook.Worksheets[i];

                currentWorksheet.Name = currentTable.TableName;

                ImportDataFromDataSetToWorksheet(currentTable, currentWorksheet);

                //SetStyle
                // SpreadSheetHelper.SetStyleSheet(this.grdVIEW, SheetIndex, currentTable.Rows.Count, currentTable.Columns.Count, 0, 0, false, false, true);

                //Only show working area 
                //SpreadSheetHelper.SetWorkingAreaVisible(this.grdVIEW, SheetIndex, currentTable.Rows.Count, currentTable.Columns.Count);


                //if (i == 0)
                //{
                //    //column width 확장
                //    ExtendWidth(SheetIndex, 0, currentTable.Columns.Count, 30);

                //    //컬럼 사용자 정렬
                //    Dictionary<int, RadHorizontalAlignment> horizonCols = new Dictionary<int, RadHorizontalAlignment>();
                //    for (int j = 0; j < dataSet.Tables[i].Columns.Count; j++)
                //    {
                //        switch (j)
                //        {
                //            case 0:
                //                horizonCols.Add(j, RadHorizontalAlignment.Left);
                //                break;
                //            case 1:
                //                horizonCols.Add(j, RadHorizontalAlignment.Center);
                //                break;
                //            case 2:
                //                horizonCols.Add(j, RadHorizontalAlignment.Right);
                //                break;
                //        }
                //    }
                //    SetContentsHorizontalCols(SheetIndex, dataSet.Tables[i].Rows.Count, horizonCols);
                //}
                //else if (i == 1)
                //{


                //    Dictionary<int, double> colWidths = new Dictionary<int, double>();
                //    int colIdx = 0;
                //    double colWith = 0;
                //    for (int j = 0; j < dataSet.Tables[i].Columns.Count;j++)
                //    {
                //        colIdx = j;
                //        switch (j)
                //        {
                //            case 0:
                //                colWith = 50;
                //                colWidths.Add(colIdx, colWith);
                //                break;
                //            case 1:
                //                colWith = 100;
                //                colWidths.Add(colIdx, colWith);
                //                break;
                //            case 2:
                //                colWith = 70;
                //                colWidths.Add(colIdx, colWith);
                //                break;
                //            case 3:
                //                colWith = 100;
                //                colWidths.Add(colIdx, colWith);
                //                break;
                //        }
                //    }
                //    SetWitth(SheetIndex, colWidths);

                //    double[] columnWidths = new double[] { 60, 150, 90, 80 };
                //    SetWitth(SheetIndex, columnWidths);
                //}
            }

            /* Excel Import */
            //Workbook wb = LoadWorkbook("0기타업체미결.xlsx", "01");
            //if(wb != null)
            //    radSpreadsheet.Workbook = wb;

            //Scroll
            //SetVertical();

            //AddHyperLink : url은 않됨.
            //AddLink(1, 0, 3, "https://www.telerik.com/forums/how-can-i-programmatically-insert-a-hyperlink-into-a-cell", "Click for more about");

            //            If you want to get the formula expression string like "=1+2" you have to use
            //string formulaString = worksheet.Cells[0, 0].GetValue().Value.RawValue;
            //            If you want to get the calculated formula result value as string you have to use:
            //worksheet.Cells[0, 0].GetValue().Value.GetResultValueAsString(CellValueFormat.GeneralFormat);
            //            If you want to get the formula value as FormulaCellValue object you have to use:
            //            FormulaCellValue formulaValue = worksheet.Cells[0, 0].GetValue().Value as FormulaCellValue;

            //grdVIEW.Workbook.Worksheets[1].Cells[0, 0, 3,3].SetIsLocked(false);
            //grdVIEW.Workbook.Worksheets[1].Protect("", WorksheetProtectionOptions.Default);
        }

        //private void SetContentsHorizontalCols(int sheetIndex, int RowCount, Dictionary<int, RadHorizontalAlignment> horizonCols)
        //{
        //    foreach (int key in horizonCols.Keys)
        //    {
        //        //grdVIEW.Workbook.Worksheets[sheetIndex].Columns[key].SetWidth(new ColumnWidth(colWith, true));

        //        RadHorizontalAlignment newAlign;
        //        if (horizonCols.TryGetValue(key, out newAlign))
        //        {
        //            //grdVIEW.Workbook.Worksheets[sheetIndex].Columns[key].SetHorizontalAlignment(newAlign);
        //            CellSelection cells = grdVIEW.Workbook.Worksheets[sheetIndex].Cells[1, key, RowCount, horizonCols.Count - 1];
        //            cells.SetHorizontalAlignment(newAlign);
        //        }
        //    }
        //}

        //private void SetWitth(int sheetIndex, double[] columnWidths)
        //{
        //    int idx = 0;
        //    foreach(double dbl in columnWidths)
        //    {
        //        grdVIEW.Workbook.Worksheets[sheetIndex].Columns[idx].SetWidth(new ColumnWidth(dbl, true));
        //        idx++;
        //    }
        //}

        ////Column With
        //private void SetWitth(int sheetIndex, Dictionary<int, double> colWidths)
        //{
        //    double colWith = 0;
        //    foreach(int key in colWidths.Keys)
        //    {
        //        if(colWidths.TryGetValue(key, out colWith))
        //        {
        //            grdVIEW.Workbook.Worksheets[sheetIndex].Columns[key].SetWidth(new ColumnWidth(colWith, true));
        //        }

        //    }
        //}

        //private void ExtendWidth(int sheetIndex, int colStart, int colEnd, double dblValue)
        //{
        //    double dblCurrentColWith = 0;
        //    for (int i = sheetIndex; i < colEnd; i++)
        //    {
        //        if(double.TryParse(grdVIEW.Workbook.Worksheets[sheetIndex].Columns[i].GetWidth().Value.Value.ToString(), out dblCurrentColWith))
        //        {
        //            dblCurrentColWith += dblValue;
        //            grdVIEW.Workbook.Worksheets[sheetIndex].Columns[i].SetWidth(new ColumnWidth(dblCurrentColWith, true));
        //        }
        //    }
        //}

        //private void AddLink(int intCurrentIdx, int intRowCnt, int intColCnt, string strLink, string titile)
        //{
        //    HyperlinkInfo link = HyperlinkInfo.CreateInDocumentHyperlink(strLink.ToString(), titile.ToString());
        //    grdVIEW.Workbook.Worksheets[intCurrentIdx].Hyperlinks.Add(new CellIndex(intRowCnt, intColCnt), link);
        //}

        ////https://www.telerik.com/forums/spreadsheet-scroll-mode
        //private void SetVertical()
        //{
        //    RadWorksheetEditor radWorksheet = (RadWorksheetEditor)this.grdVIEW.ActiveSheetEditor;
        //    if(radWorksheet!=null)
        //        radWorksheet.VerticalScrollMode = ScrollMode.PixelBased;
        //}


        ////Style
        //private void SetStyle(int intCurrentIdx, int intRowCnt, int intColCnt, int intFixRow=0, int intFixCol = 0)
        //{
        //    grdVIEW.Workbook.Worksheets[SheetIndex].Rows[0].SetVerticalAlignment(RadVerticalAlignment.Center); //해더 세로 중간 정렬

        //    //Column Autofit
        //    for (int i=0; i<intColCnt; i++)
        //    {
        //        grdVIEW.Workbook.Worksheets[SheetIndex].Columns[i].SetVerticalAlignment(RadVerticalAlignment.Center); //세로 중간
        //        grdVIEW.Workbook.Worksheets[SheetIndex].Columns[i].AutoFitWidth(); //Autofit
        //        grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[0, i].SetHorizontalAlignment(RadHorizontalAlignment.Center);//Header만
        //        //grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[0, i]
        //    }

        //    //Header
        //    CellSelection headerSelection = grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[0, 0, 0, intColCnt-1];
        //    ThemableColor gray = new ThemableColor(Colors.LightGray);
        //    ThemableColor patternColor = new ThemableColor(Colors.LightGray);
        //    ThemableColor backgroundColor = new ThemableColor(Colors.LightGray);
        //    IFill fill = new PatternFill(PatternType.Gray75Percent, patternColor, backgroundColor);
        //    headerSelection.SetFill(fill); //header color

        //    //Row
        //    CellSelection rowSelection = grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[1, 1, intRowCnt-1, intColCnt-1];

        //    //All
        //    SetBorders(intCurrentIdx, 0, 0, intRowCnt, intColCnt - 1, Colors.Black);

        //    //Freeze
        //    if (intFixRow>0 || intFixCol>0)
        //    {
        //        RadWorksheetEditor worksheetEditor = grdVIEW.ActiveWorksheetEditor;
        //        worksheetEditor.FreezePanes(new CellIndex(intFixRow, intFixCol));
        //    }

        //}

        ////Fore Color
        //private void SetForeColor(int intCurrentIdx, int intRowCnt, int intColCnt, Color color)
        //{
        //    grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[intRowCnt, intColCnt].SetForeColor(new ThemableColor(color));
        //}
        //private void SetForeColor(int intCurrentIdx, int intStartRowCnt, int intStartColCnt, int intRowCnt, int intColCnt, Color color)
        //{
        //    grdVIEW.Workbook.Worksheets[intCurrentIdx].Cells[intStartRowCnt, intStartColCnt, intRowCnt, intColCnt].SetForeColor(new ThemableColor(color));
        //}
        ////Back Color
        //private void SetBackColor(int intCurrentIdx, int intRowCnt, int intColCnt, Color color)
        //{
        //    CellSelection selCols = grdVIEW.Workbook.Worksheets[SheetIndex].Cells[intRowCnt, intColCnt];
        //    ThemableColor gray = new ThemableColor(Colors.LightGray);
        //    ThemableColor patternColor = new ThemableColor(Colors.LightGray);
        //    ThemableColor backgroundColor = new ThemableColor(Colors.LightGray);
        //    IFill fill = new PatternFill(PatternType.Gray75Percent, patternColor, backgroundColor);
        //    selCols.SetFill(fill);
        //}
        //private void SetBackColor(int intCurrentIdx, int intStartRowCnt, int intStartColCnt, int intRowCnt, int intColCnt, Color color)
        //{
        //    CellSelection selCols = grdVIEW.Workbook.Worksheets[SheetIndex].Cells[intStartRowCnt, intStartColCnt, intRowCnt, intColCnt];
        //    ThemableColor gray = new ThemableColor(color);
        //    ThemableColor patternColor = new ThemableColor(color);
        //    ThemableColor backgroundColor = new ThemableColor(color);
        //    IFill fill = new PatternFill(PatternType.Gray75Percent, patternColor, backgroundColor);
        //    selCols.SetFill(fill);
        //}

        ////Border
        //private void SetBorders(int intCurrentIdx, int intRowCnt, int intColCnt, Color color)
        //{
        //    CellSelection selCols = grdVIEW.Workbook.Worksheets[SheetIndex].Cells[intRowCnt, intColCnt];
        //    ThemableColor selColor = new ThemableColor(color);
        //    CellBorders blackBorders = new CellBorders(
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.None, selColor),
        //        new CellBorder(CellBorderStyle.None, selColor),
        //        new CellBorder(CellBorderStyle.None, selColor),
        //        new CellBorder(CellBorderStyle.None, selColor)
        //        );
        //    selCols.SetBorders(blackBorders);
        //}
        //private void SetBorders(int intCurrentIdx, int intStartRowCnt, int intStartColCnt, int intRowCnt, int intColCnt, Color color)
        //{
        //    CellSelection selCols = grdVIEW.Workbook.Worksheets[SheetIndex].Cells[intStartRowCnt, intStartRowCnt, intRowCnt, intColCnt];
        //    ThemableColor selColor = new ThemableColor(color);
        //    CellBorders blackBorders = new CellBorders(
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.None, selColor),
        //        new CellBorder(CellBorderStyle.None, selColor)
        //        );
        //    selCols.SetBorders(blackBorders);
        //}
        //private void SetBorder(int intCurrentIdx, int intStartRowCnt, int intStartColCnt, int intRowCnt, int intColCnt, Color color)
        //{
        //    CellSelection selCols = grdVIEW.Workbook.Worksheets[SheetIndex].Cells[intStartRowCnt, intStartColCnt, intRowCnt, intColCnt];
        //    ThemableColor selColor = new ThemableColor(color);
        //    CellBorders blackBorders = new CellBorders(
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.Thin, selColor),
        //        new CellBorder(CellBorderStyle.None, selColor),
        //        new CellBorder(CellBorderStyle.None, selColor),
        //        new CellBorder(CellBorderStyle.None, selColor),
        //        new CellBorder(CellBorderStyle.None, selColor)
        //        );
        //    selCols.SetBorders(blackBorders);
        //}
        //엑셀 로드
        private Workbook LoadWorkbook(string FilePath, string type)
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
        //Table input
        private void ImportDataFromDataSetToWorksheet(DataTable table, Worksheet worksheet)
        {
            for (int col = 0; col < table.Columns.Count; col++)
            {
                CellSelection cell = worksheet.Cells[0, col];
                cell.SetValue(table.Columns[col].ToString());
            }

            for (int row = 0; row < table.Rows.Count; row++)
            {
                var currentRow = table.Rows[row];

                for (int col = 0; col < table.Columns.Count; col++)
                {
                    CellSelection cell = worksheet.Cells[row + 1, col];
                    var currentCellValue = currentRow.ItemArray[col];

                    if (currentCellValue.GetType() == typeof(double) || currentCellValue.GetType() == typeof(int))
                    {
                        double _dbl = 0;
                        if (double.TryParse(currentCellValue.ToString(), out _dbl))
                            cell.SetValue(_dbl);
                        else
                            cell.SetValue(currentCellValue.ToString());
                    }
                    else if (currentCellValue.GetType() == typeof(DateTime))
                    {
                        cell.SetValue((DateTime)currentCellValue);
                    }
                    else if (currentCellValue.GetType() == typeof(bool))
                    {
                        cell.SetValue((bool)currentCellValue);
                    }
                    else
                    {
                        cell.SetValue(currentCellValue.ToString());
                    }
                }
            }
        }
        //Temp DataSet
        private DataSet InitDs()
        {
            DataSet _Ds = new DataSet();
            DataTable dt = new DataTable();
            dt.TableName = "dt1";
            DataColumn col1 = new DataColumn();
            col1.DataType = System.Type.GetType("System.Int32");
            col1.ColumnName = "ID";
            col1.AutoIncrement = true;
            dt.Columns.Add(col1);

            DataColumn col2 = new DataColumn();
            col2.DataType = System.Type.GetType("System.String");
            col2.ColumnName = "Name";
            col2.DefaultValue = "No Name";
            dt.Columns.Add(col2);

            DataColumn col3 = new DataColumn();
            col3.DataType = System.Type.GetType("System.Int32");
            col3.ColumnName = "Age";
            col3.DefaultValue = 0;
            dt.Columns.Add(col3);
            DataRow row = dt.NewRow();
            row["ID"] = 1;
            row["Name"] = "Jone";
            row["Age"] = 33;
            dt.Rows.Add(row);
            row = dt.NewRow();
            row["ID"] = 2;
            row["Name"] = "Indocin777";
            row["Age"] = 13;
            row["ID"] = 3;
            row["Name"] = "Nana";
            row["Age"] = 22;
            dt.Rows.Add(row);

            _Ds.Tables.Add(dt);

            DataTable table = new DataTable();
            table.TableName = "dt2";
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            // Step 3: here we add 5 rows.
            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);
            for (int i = 0; i < 10; i++)
            {
                table.Rows.Add(250 + i, "Hydralazine" + i, "David" + i, DateTime.Now);
            }
            table.Rows.Add(26, "Combivent", "Janet2", DateTime.Now);
            table.Rows.Add(221, "Hydralazine", "Janet1", DateTime.Now);
            table.Rows.Add(213, "Combivent", "Janet2", DateTime.Now);
            _Ds.Tables.Add(table);
            return _Ds;
        }
        #endregion
    }
}

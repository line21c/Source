using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Telerik.Windows.Controls;
using Telerik.Windows.Documents.Spreadsheet.Model;
//https://www.telerik.com/forums/radspreadsheet-issue-if-used-inside-a-datatemplate

namespace TelerikWpf.Test.Helpers
{
    public static class SpreadsheetExtensions
    {
        public static Workbook GetBindableWorkbook(DependencyObject obj)
        {
            return (Workbook)obj.GetValue(BindableWorkbookProperty);
        }

        public static void SetBindableWorkbook(DependencyObject obj, Workbook value)
        {
            obj.SetValue(BindableWorkbookProperty, value);
        }

        // Using a DependencyProperty as the backing store for Workbook.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty BindableWorkbookProperty =
            DependencyProperty.RegisterAttached("BindableWorkbook", typeof(Workbook), typeof(RadSpreadsheet),
                new PropertyMetadata(OnBindableWorkbookPropertyChanged));

        private static void OnBindableWorkbookPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            RadSpreadsheet radSpreadsheet = d as RadSpreadsheet;

            if (radSpreadsheet != null)
            {
                radSpreadsheet.Workbook = e.NewValue as Workbook;
            }
        }
    }
}

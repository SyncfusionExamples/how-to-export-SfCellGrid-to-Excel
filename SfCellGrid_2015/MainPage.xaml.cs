using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Grid;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Graphics.Printing;
using Windows.Storage;
using Windows.UI;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Windows.UI.Xaml.Printing;


namespace SfCellGrid_2015
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {

        public MainPage()
        {
            this.InitializeComponent();
            CellGrid.Model.QueryCellInfo += Model_QueryCellInfo;
            
        }

        private void Model_QueryCellInfo(object sender, Syncfusion.UI.Xaml.CellGrid.Styles.GridQueryCellInfoEventArgs e)
        {
            e.Style.CellValue = string.Format("R{0}:C{1}", e.Cell.RowIndex, e.Cell.ColumnIndex);
            if (e.Cell.ColumnIndex == 2)
                e.Style.Background = new SolidColorBrush(Windows.UI.Color.FromArgb(255, 90, 90, 90));
        }

        private async void button_Click(object sender, RoutedEventArgs e)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2016;
            IApplication application = excelEngine.Excel;

            IWorkbook workbook = application.Workbooks.Create(1);

            IWorksheet sheet = workbook.Worksheets[0];

            for (int i = 0; i < CellGrid.RowCount; i++)
            {
                for (int j = 0; j < CellGrid.ColumnCount; j++)
                {
                    IRange range = sheet[i + 1, j + 1];
                    var style = CellGrid.Model[i, j];
                    var brush = (style.Background as SolidColorBrush);
                    //Export with style
                    if (brush != null)
                        range.CellStyle.Color = brush.Color;
                    range.Value = style.CellValue.ToString();
                }
            }
            StorageFile storageFile;
            StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
            storageFile = await local.CreateFileAsync("Sample.xlsx", CreationCollisionOption.ReplaceExisting);
            await workbook.SaveAsAsync(storageFile);
        }

        private async void button2_Click(object sender, RoutedEventArgs e)
        {
            //Create a new PDF document.

            PdfDocument pdfDocument = new PdfDocument();

            //Create the page

            PdfPage pdfPage = pdfDocument.Pages.Add();

            //Create the parent grid

            PdfGrid parentPdfGrid = new PdfGrid();

            //Add the rows
            for (int i = 0; i < CellGrid.RowCount; i++)
            {
                PdfGridRow row1 = parentPdfGrid.Rows.Add();
                row1.Height = 50;
                parentPdfGrid.Columns.Add(CellGrid.ColumnCount);
                for (int j = 0; j < CellGrid.ColumnCount; j++)
                {
                    var style = CellGrid.Model[i, j];
                    PdfGridCell pdfGridCell = parentPdfGrid.Rows[i].Cells[j];
                    pdfGridCell.Value = style.CellValue;
                    var brush = (style.Background as SolidColorBrush);
                    //Export with style
                    //if (brush != null)
                    //    pdfGridCell.Style.BackgroundBrush

                }
            }



            //Draw the PdfGrid.

            parentPdfGrid.Draw(pdfPage, PointF.Empty);
            StorageFile storageFile;
            StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
            storageFile = await local.CreateFileAsync("Sample.pdf", CreationCollisionOption.ReplaceExisting);
            //Save the document.

            await pdfDocument.SaveAsync(storageFile);

        }

        
    }
}

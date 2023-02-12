using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace PaymentsExampleApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private PaymentsBaseEntities _context = new PaymentsBaseEntities();
        public MainWindow()
        {
            InitializeComponent();
            ChartPayments.ChartAreas.Add(new ChartArea("Main"));

            var currentSeries = new Series("Payments")
            {
                IsValueShownAsLabel = true
            };
            ChartPayments.Series.Add(currentSeries);

            ComboUser.ItemsSource = _context.Users.ToList();
            ComboCharTypes.ItemsSource = Enum.GetValues(typeof(SeriesChartType));
        }

        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            if (ComboUser.SelectedItem is User currentUser &&
                ComboCharTypes.SelectedItem is SeriesChartType currentType)
            {
                Series currentSeries = ChartPayments.Series.FirstOrDefault();
                currentSeries.ChartType = currentType;
                currentSeries.Points.Clear();

                var categoriesList = _context.Categories.ToList();
                foreach (var category in categoriesList)
                {
                    currentSeries.Points.AddXY(category.Name,
                        _context.Payments.ToList().Where(p => p.User == currentUser
                        && p.Category == category).Sum(p => p.Price * p.Num));
                }
            }
        }

        private void BtnExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            var allUser = _context.Users.ToList().OrderBy(p => p.FIO).ToList();

            var application = new Excel.Application();
            application.SheetsInNewWorkbook = allUser.Count();

            Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);

            for (int i = 0; i < allUser.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = application.Worksheets.Item[i + 1];
                worksheet.Name = allUser[i].FIO;

                worksheet.Cells[1][startRowIndex] = "Дата платежа";
                worksheet.Cells[2][startRowIndex] = "Название";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                worksheet.Cells[4][startRowIndex] = "Количество";
                worksheet.Cells[5][startRowIndex] = "Сумма";

                startRowIndex++;

                var usersCategories = allUser[i].Payments.OrderBy(p => p.Date).GroupBy(p => p.Category).OrderBy(p => p.Key.Name);

                foreach (var groupCategory in usersCategories)
                {
                    Excel.Range headerRande = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[5][startRowIndex]];
                    headerRande.Merge();
                    headerRande.Value = groupCategory.Key.Name;
                    headerRande.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    headerRande.Font.Italic = true;

                    startRowIndex++;

                    foreach (var payment in groupCategory)
                    {
                        worksheet.Cells[1][startRowIndex] = payment.Date.ToString("dd.MM.yyyy HH:mm");
                        worksheet.Cells[2][startRowIndex] = payment.Name;
                        worksheet.Cells[3][startRowIndex] = payment.Price;
                        worksheet.Cells[4][startRowIndex] = payment.Num;

                        worksheet.Cells[5][startRowIndex].Formula = $"=C{startRowIndex}*D{startRowIndex}";

                        worksheet.Cells[3][startRowIndex].NumberFormat = "####,00";

                        startRowIndex++;
                    }
                    Excel.Range sumRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[4][startRowIndex]];
                    sumRange.Merge();
                    sumRange.Value = "ИТОГО:";
                    sumRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

                    worksheet.Cells[5][startRowIndex].Formula = $"=SUM(E{startRowIndex - groupCategory.Count()}:" +
                        $"E{startRowIndex - 1})";

                    sumRange.Font.Bold = worksheet.Cells[5][startRowIndex].Font.Bold = true;

                    worksheet.Cells[5][startRowIndex].NumberFormat = "####,00";


                    startRowIndex++;

                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex - 1]];
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                    worksheet.Columns.AutoFit();
                }
            }
            application.Visible = true;
        }

        private void BtnExportToWord_Click(object sender, RoutedEventArgs e)
        {
            var allUser = _context.Users.ToList();
            var allCategories = _context.Categories.ToList();

            var application = new Word.Application();

            Word.Document document = application.Documents.Add();

            foreach (var user in allUser)
            {
                Word.Paragraph userParagraph = document.Paragraphs.Add();
                Word.Range userRange = userParagraph.Range;
                userRange.Text = user.FIO;
                userParagraph.set_Style("Заголовок");
                userRange.InsertParagraphAfter();

                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 3);
                paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                Word.Range cellRange;

                cellRange = paymentsTable.Cell(1, 1).Range;
                cellRange.Text = "Иконка";
                cellRange = paymentsTable.Cell(1, 2).Range;
                cellRange.Text = "Категория";
                cellRange = paymentsTable.Cell(1, 3).Range;
                cellRange.Text = "Сумма расходов";

                paymentsTable.Rows[1].Range.Bold = 1;
                paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                for (int i = 0; i < allCategories.Count(); i++)
                {
                    var currentCategory = allCategories[i];
                    var path = AppDomain.CurrentDomain.BaseDirectory;

                    cellRange = paymentsTable.Cell(i + 2, 1).Range;
                    Word.InlineShape imageShape = cellRange.InlineShapes.AddPicture(path.Substring(0, path.Length - 10)
                        + @"\Assets\" + currentCategory.Icon);
                    imageShape.Width = imageShape.Height = 40;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = paymentsTable.Cell(i + 2, 2).Range;
                    cellRange.Text = currentCategory.Name;

                    cellRange = paymentsTable.Cell(i + 2, 3).Range;
                    cellRange.Text = user.Payments.ToList()
                        .Where(p => p.Category == currentCategory).Sum(p => p.Num * p.Price).ToString("N2") + " руб.";
                }

                Payment maxPayment = user.Payments
                    .OrderByDescending(p => p.Price * p.Num).FirstOrDefault();
                if (maxPayment != null)
                {
                    Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                    Word.Range maxPaymentRange = maxPaymentParagraph.Range;
                    maxPaymentRange.Text = $"Самый дорогостоящий платеж - {maxPayment.Name} за {(maxPayment.Price * maxPayment.Num).ToString("N2")}"
                        + $"руб. от {maxPayment.Date.ToString("dd.MM.yyyy HH:mm")}";
                    maxPaymentParagraph.set_Style("Выделенная цитата");
                    maxPaymentRange.Font.Color = Word.WdColor.wdColorDarkRed;
                    maxPaymentRange.InsertParagraphAfter();
                }
                Payment minPayment = user.Payments
                    .OrderBy(p => p.Price * p.Num).FirstOrDefault();
                if (minPayment != null)
                {
                    Word.Paragraph minPaymentParagraph = document.Paragraphs.Add();
                    Word.Range minPaymentRange = minPaymentParagraph.Range;
                    minPaymentRange.Text = $"Самый дешевый платеж - {minPayment.Name} за {(minPayment.Price * minPayment.Num).ToString("N2")}"
                        + $"руб. от {minPayment.Date.ToString("dd.MM.yyyy HH:mm")}";
                    minPaymentParagraph.set_Style("Выделенная цитата");
                    minPaymentRange.Font.Color = Word.WdColor.wdColorDarkGreen;
                }

                if (user != allUser.LastOrDefault())
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
            }

            application.Visible = true;

            document.SaveAs2(@"C:\Test.docx");
            document.SaveAs2(@"C:\Test.pdf", Word.WdExportFormat.wdExportFormatPDF);
        }
    }
}

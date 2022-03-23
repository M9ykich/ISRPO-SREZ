using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using ScottPlot;
using ScottPlot.Plottable;
using AppSrez2.models;



namespace AppSrez2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        
        public MainWindow()
        {
            InitializeComponent();
        }
        public static List<Sale> clients = null;
        /// <summary>
        /// Метод для выгрузки данных по датам
        /// </summary>
        private async void BtnGet_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (HttpClient httpClient = new HttpClient { BaseAddress = new Uri(Properties.Settings.Default.BaseAddress) })
                {
                    var content = new StringContent("", Encoding.UTF8, "application/json");
                    HttpResponseMessage httpResponseMessage = await httpClient.PostAsync($"api/Sale?dateStart={dateStart.SelectedDate.Value.Date.ToString("yyyy-MM-dd")}&dateEnd={dateEnd.SelectedDate.Value.Date.ToString("yyyy-MM-dd")}", content);
                    string data = httpResponseMessage.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                    clients = JsonSerializer.Deserialize<List<Sale>>(data);
                    DgSale.ItemsSource = clients;
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Ошибка");
            }
           
        }
        /// <summary>
        /// Метод для формирования чека в ворд
        /// </summary>
        private void BtnWord_Click(object sender, RoutedEventArgs e)
        {
                try
                {
                    var a = DgSale.SelectedItem as Sale;
                    Word._Application wApp = new Word.Application(); 
                    Word._Document wDoc = wApp.Documents.Add();
                    wApp.Visible = false;
                    wDoc.Activate();
                    float cost = 0;
                    var ClientParagraph = wDoc.Content.Paragraphs.Add();
                    ClientParagraph.Range.Text = $"Фамилия:\t{a.client.lastName}\n" +
                        $"Имя:\t{a.client.firstName}\n" +
                        $"Отчество:\t{a.client.patronymic}\n";
                    Word.Table wTable = wDoc.Tables.Add((Word.Range)ClientParagraph.Range,
                        a.telephones.Length + 1, 6, Word.WdDefaultTableBehavior.wdWord9TableBehavior);
                    wTable.Cell(1, 1).Range.Text = "Артикул";
                    wTable.Cell(1, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 2).Range.Text = "Наименование";
                    wTable.Cell(1, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 3).Range.Text = "Категория";
                    wTable.Cell(1, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 4).Range.Text = "Количество";
                    wTable.Cell(1, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 5).Range.Text = "Цена";
                    wTable.Cell(1, 5).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 6).Range.Text = "Манафактура";
                    wTable.Cell(1, 6).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    int countRow = 2;
                    foreach (var item in a.telephones)
                    {
                        wTable.Cell(countRow, 1).Range.Text = item.articul.ToString();
                        wTable.Cell(countRow, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 2).Range.Text = item.nameTelephone.ToString();
                        wTable.Cell(countRow, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 3).Range.Text = item.category.ToString();
                        wTable.Cell(countRow, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 4).Range.Text = item.count.ToString();
                        wTable.Cell(countRow, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 5).Range.Text = item.cost.ToString();
                        wTable.Cell(countRow, 5).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 6).Range.Text = item.manufacturer.ToString();
                        wTable.Cell(countRow, 6).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cost += item.cost * item.count;
                        countRow++;
                    }
                    var CostParagraph = wDoc.Content.Paragraphs.Add();
                    CostParagraph.Range.Text = $"Стоимость:\t{cost}\n";
                    wDoc.SaveAs2($@"{Environment.CurrentDirectory}\1.docx");
                    wDoc.SaveAs2($@"{Environment.CurrentDirectory}\1.pdf", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                    MessageBox.Show("Чек сформирован");
                    wDoc.Close();
                 }
                catch (Exception ex)
                {

                    MessageBox.Show($"Ошибка!");
                }
        }
        /// <summary>
        /// Метод для формирования чека в Excel
        /// </summary>
        private void BtnExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var a = DgSale.SelectedItem as Sale;
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.SheetsInNewWorkbook = 2;
                Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
                excelApp.DisplayAlerts = false;
                Excel.Worksheet sheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(1);
                sheet.Name = "Чек";
                sheet.Columns[2].ColumnWidth = 15;
                sheet.Cells[1, 1] = "Артикул";
                sheet.Cells[1, 2] = "Наименование";
                sheet.Cells[1, 3] = "Категория";
                sheet.Cells[1, 4] = "Количество";
                sheet.Cells[1, 5] = "Цена";
                sheet.Cells[1, 6] = "Манафактура";
                int countrow = 2;
                foreach (var item in a.telephones)
                {
                    sheet.Cells[countrow, 1] = item.articul;
                    sheet.Cells[countrow, 2] = item.nameTelephone;
                    sheet.Cells[countrow, 3] = item.category;
                    sheet.Cells[countrow, 4] = item.count;
                    sheet.Cells[countrow, 5] = item.cost;
                    sheet.Cells[countrow, 6] = item.manufacturer;
                    sheet.Cells[countrow, 7].Formula = $"=D{countrow}*E{countrow}";
                    countrow++;
                }
                sheet.Cells[1, 8] = "Итого";
                sheet.Cells[2, 8].Formula = $"=SUM(G2:G{countrow - 1} )";
                sheet.SaveAs($@"{ Environment.CurrentDirectory}\1.xlsx");
                MessageBox.Show("Файл сохранен!");
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка!");
            }
        }
        /// <summary>
        /// Метод для создания отчетности в Word
        /// </summary>
        private void BtnWord1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Word._Application wApp = new Word.Application();
                Word._Document wDoc = wApp.Documents.Add();
                wApp.Visible = false;
                wDoc.Activate();
                foreach (var item in clients)
                {
                    float cost = 0;
                    var ClientParagraph = wDoc.Content.Paragraphs.Add();
                    ClientParagraph.Range.Text = $"Фамилия:\t{item.client.lastName}\n" +
                        $"Имя:\t{item.client.firstName}\n" +
                        $"Отчество:\t{item.client.patronymic}\n";
                    Word.Table wTable = wDoc.Tables.Add((Word.Range)ClientParagraph.Range,
                        item.telephones.Length + 1, 6, Word.WdDefaultTableBehavior.wdWord9TableBehavior);
                    wTable.Cell(1, 1).Range.Text = "Артикул";
                    wTable.Cell(1, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 2).Range.Text = "Наименование";
                    wTable.Cell(1, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 3).Range.Text = "Категория";
                    wTable.Cell(1, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 4).Range.Text = "Количество";
                    wTable.Cell(1, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 5).Range.Text = "Цена";
                    wTable.Cell(1, 5).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    wTable.Cell(1, 6).Range.Text = "Манафактура";
                    wTable.Cell(1, 6).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    int countRow = 2;
                    foreach (var telephone in item.telephones)
                    {
                        wTable.Cell(countRow, 1).Range.Text = telephone.articul.ToString();
                        wTable.Cell(countRow, 1).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 2).Range.Text = telephone.nameTelephone.ToString();
                        wTable.Cell(countRow, 2).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 3).Range.Text = telephone.category.ToString();
                        wTable.Cell(countRow, 3).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 4).Range.Text = telephone.count.ToString();
                        wTable.Cell(countRow, 4).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 5).Range.Text = telephone.cost.ToString();
                        wTable.Cell(countRow, 5).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        wTable.Cell(countRow, 6).Range.Text = telephone.manufacturer.ToString();
                        wTable.Cell(countRow, 6).Range.Paragraphs.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cost += telephone.cost * telephone.count;
                        countRow++;
                    }
                    var CostParagraph = wDoc.Content.Paragraphs.Add();
                    CostParagraph.Range.Text = $"Стоимость:\t{cost}\n";
                }
                wDoc.SaveAs2($@"{Environment.CurrentDirectory}\all2.docx");
                wDoc.SaveAs2($@"{Environment.CurrentDirectory}\all2.pdf", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                MessageBox.Show("Чек сформирован");
                wDoc.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show($"Ошибка!");
            }
        }
        /// <summary>
        /// Метод для формирования отчетности в Excel 
        /// </summary>
        private void BtnExcel1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;
                int count = 0;
                foreach (var item in clients)
                {
                    excelApp.SheetsInNewWorkbook = 2;
                    count++;
                    Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);
                    excelApp.DisplayAlerts = false;
                    Excel.Worksheet sheet = (Excel.Worksheet)excelApp.Worksheets.get_Item(2);
                    sheet.Name = "Чек";
                    sheet.Columns[2].ColumnWidth = 15;
                    sheet.Cells[1, 1] = "Артикул";
                    sheet.Cells[1, 2] = "Наименование";
                    sheet.Cells[1, 3] = "Категория";
                    sheet.Cells[1, 4] = "Количество";
                    sheet.Cells[1, 5] = "Цена";
                    sheet.Cells[1, 6] = "Манафактура";
                    int countrow = 2;
                    foreach (var telephone in item.telephones)
                    {
                        sheet.Cells[countrow, 1] = telephone.articul;
                        sheet.Cells[countrow, 2] = telephone.nameTelephone;
                        sheet.Cells[countrow, 3] = telephone.category;
                        sheet.Cells[countrow, 4] = telephone.count;
                        sheet.Cells[countrow, 5] = telephone.cost;
                        sheet.Cells[countrow, 6] = telephone.manufacturer;
                        sheet.Cells[countrow, 7].Formula = $"=D{countrow}*E{countrow}";
                        countrow++;
                    }
                    sheet.Cells[1, 8] = "Итого";
                    sheet.Cells[2, 8].Formula = $"=SUM(G2:G{countrow - 1} )";
                    sheet.SaveAs($@"{ Environment.CurrentDirectory}\all{count}.xlsx");
                }
                MessageBox.Show("Файл сохранен!");
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка!");
            }
        }
        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (ComboBox1.SelectedIndex == 0)
                {
                    StackP1.Visibility = Visibility.Visible;
                    StackP2.Visibility = Visibility.Collapsed;
                    AppSales();
                }
                else if (ComboBox1.SelectedIndex == 1)
                {
                    StackP2.Visibility = Visibility.Visible;
                    StackP1.Visibility = Visibility.Collapsed;
                    AppDate();
                }
            }
            catch (Exception)
            {

                MessageBox.Show("Ошибка!");
            }
            
        }
        /// <summary>
        /// Метод для вывода данных о продаже
        /// </summary>
        public void AppSales()
        {
            try
            {
                if(clients == null)
                {
                    MessageBox.Show("Ошибка!");
                }
                else
                {
                    List<string> name = new List<string>();
                    List<double> countsale = new List<double>();
                    int count = 0;
                    foreach (var item in clients)
                    {

                        foreach (var telephone in item.telephones)
                        {
                            if (!name.Contains(telephone.manufacturer))
                            {
                                name.Add(telephone.manufacturer);
                                countsale.Add(telephone.count * telephone.cost);
                            }
                            else
                            {
                                var x = countsale.ElementAt(name.IndexOf(telephone.manufacturer));
                                x += telephone.count * telephone.cost;
                            }
                        }

                    }
                    var pie = WpfPlot.Plot.AddPie(countsale.ToArray());
                    pie.SliceLabels = name.ToArray();
                    pie.ShowPercentages = true;
                    pie.ShowValues = true;
                    pie.ShowLabels = true;
                    WpfPlot.Plot.Legend();
                    WpfPlot.Refresh();
                }
               
                
            }
            catch (Exception)
            {

                MessageBox.Show("Ошибка!");
            }
           
        }
        /// <summary>
        /// Метод для вывода графика по дате
        /// </summary>
        public void AppDate()
        {
            try
            {
                int count = (dateEnd.SelectedDate.Value.Date - dateStart.SelectedDate.Value.Date).Days + 1;
                double[] countDate = new double[count];
                DateTime[] dates = new DateTime[count];
                dates[0] = dateStart.SelectedDate.Value.Date;

                for (int i = 0; i < count; i++)
                {
                    dates[i] = dates[0].AddDays(i);
                    countDate[i] = clients.Where(c => c.dateSale == dates[i]).Sum(x => x.telephones.Sum(c => c.count * c.cost));
                }
                double[] xs = dates.Select(x => x.ToOADate()).ToArray();
                WpfPlot1.Plot.AddScatter(xs, countDate);
                WpfPlot1.Plot.XAxis.DateTimeFormat(true);
                WpfPlot1.Refresh();
            }
            catch (Exception)
            {

                MessageBox.Show("Ошибка!");
            }
            
        }
      
    }
}

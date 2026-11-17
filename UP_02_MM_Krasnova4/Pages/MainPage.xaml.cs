using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
using UP_02_MM_Krasnova4.Classes;
using Excel = Microsoft.Office.Interop.Excel;

namespace UP_02_MM_Krasnova4.Pages
{
    /// <summary>
    /// Логика взаимодействия для MainPage.xaml
    /// </summary>
    public partial class MainPage : Page
    {
        public MainPage()
        {
            InitializeComponent();

            ModelSMOCB.ItemsSource = new List<string>
            {
                "Одноканальная СМО с отказами в обслуживании",
                "Одноканальная СМО с ограниченной длиной очереди",
                "Одноканальная СМО с неограниченной очередью",
                "Многоканальная СМО с отказами в обслуживании",
                "Многоканальная СМО с ограниченной длиной очереди",
                "Многоканальная СМО с неограниченной очередью",
            };
            ModelSMOCB.SelectedIndex = 0;

            lambdaCB.ItemsSource = new List<string> { "час", "мин", "сутки", };
            lambdaCB.SelectedIndex = 1;

            tobsCB.ItemsSource = new List<string> { "час", "мин", "сутки", };
            tobsCB.SelectedIndex = 1;
        }

        private void CleanAll()
        {
            nTB.Text = "2";
            mLenTB.Text = "1";
            mAmountTB.Text = "0";
            lambdaTB.Text = "1";
            lambdaCB.SelectedIndex = 1;
            muTB.Text = "";
            tobsTB.Text = "";
            tobsCB.SelectedIndex = 1;
            ResultSP.Visibility = Visibility.Collapsed;
        }

        private void ModelSMOCB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ResultSP.Visibility = Visibility.Collapsed;

            if (ModelSMOCB.SelectedItem == null) { return; }

            int index = ModelSMOCB.SelectedIndex;
            nSP.Visibility = (index >= 3) ? Visibility.Visible : Visibility.Collapsed;
            mLenSP.Visibility = (index == 1 || index == 4) ? Visibility.Visible : Visibility.Collapsed;
            mAmountSP.Visibility = (index == 5) ? Visibility.Visible : Visibility.Collapsed;
        }

        private double tobsConvert(double tobs)
        {
            double tobsInHours;

            switch (tobsCB.SelectedItem?.ToString())
            {
                case "мин":
                    tobsInHours = tobs / 60.0;
                    break;
                case "сутки":
                    tobsInHours = tobs * 24.0;
                    break;
                default:
                    tobsInHours = tobs;
                    break;
            }

            switch (lambdaCB.SelectedItem?.ToString())
            {
                case "мин":
                    return tobsInHours * 60.0;
                case "сутки":
                    return tobsInHours / 24.0;
                default:
                    return tobsInHours;
            }
        }

        private void ResultBtn_Click(object sender, RoutedEventArgs e)
        {
            ResultTB.Text = "";
            double lambda, mu;
            double mLen = 0;
            double mAmount = 0;
            double tobs = 1;
            int n = 1;

            if (ModelSMOCB.SelectedItem == null) { MessageBox.Show("Выберите модель СМО!"); return; }

            // проверка заполненности основных полей
            if (!double.TryParse(lambdaTB.Text, out lambda) || lambda <= 0)
            {
                MessageBox.Show("Введите корректное положительное значение λ!");
                return;
            }

            if (!string.IsNullOrWhiteSpace(muTB.Text))
            {
                if (!double.TryParse(muTB.Text, out mu) || mu <= 0)
                {
                    MessageBox.Show("Введите корректное положительное значение μ.");
                    return;
                }
            }
            else if (!string.IsNullOrWhiteSpace(tobsTB.Text))
            {
                if (!double.TryParse(tobsTB.Text, out tobs) || tobs <= 0)
                {
                    MessageBox.Show("Введите корректное положительное значение t_обс.");
                    return;
                }

                tobs = tobsConvert(tobs);
                mu = 1.0 / tobs;
            }
            else { MessageBox.Show("Заполните μ или t_обс."); return; }

            // проверка заполненности дополнительных полей
            int index = ModelSMOCB.SelectedIndex;
            if ((index >= 3 && !int.TryParse(nTB.Text, out n)) ||
               ((index == 1 || index == 4) && !double.TryParse(mLenTB.Text, out mLen)) ||
               (index == 5 && !double.TryParse(mAmountTB.Text, out mAmount))
               ) { MessageBox.Show("Корректно заполните необходимые поля!"); return; }

            string time = lambdaCB.SelectedItem.ToString();
            if (index >= 3) { ResultTB.Text += $"n = {n} канал.\n"; }
            ResultTB.Text += $"λ = {lambda} заяв/{time}\nμ = {mu} заяв/{time}\ntобс = {Math.Round(tobs, 3)} заяв/{time}\n-------------------------\n";

            // выбор метода вычисления
            switch (index)
            {
                case 0:
                    ResultTB.Text += CalcCMO.CalcSingleReject(lambda, mu);
                    break;
                case 1:
                    ResultTB.Text += CalcCMO.CalcSingleLimitedQueue(lambda, mu, mLen);
                    break;
                case 2:
                    ResultTB.Text += CalcCMO.CalcSingleInfiniteQueue(lambda, mu);
                    break;
                case 3:
                    ResultTB.Text += CalcCMO.CalcMultiReject(lambda, mu, n);
                    break;
                case 4:
                    ResultTB.Text += CalcCMO.CalcMultiLimitedQueue(lambda, mu, n, mLen);
                    break;
                case 5:
                    ResultTB.Text += CalcCMO.CalcMultiInfiniteQueue(lambda, mu, n, mAmount);
                    break;
                default: break;
            }
            ResultSP.Visibility = Visibility.Visible;
        }

        private void CleanBtn_Click(object sender, RoutedEventArgs e)
        {
            CleanAll();
        }

        private void ExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            var result = ResultTB.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            var application = new Excel.Application();
            application.SheetsInNewWorkbook = 1;
            Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);

            Excel.Worksheet worksheet = application.Worksheets.Item[1];
            worksheet.Name = "Лист";

            int startRowIndex = 1;
            worksheet.Cells[1][startRowIndex] = "Переменная";
            worksheet.Cells[2][startRowIndex] = "Значение";
            worksheet.Cells[3][startRowIndex] = "Пояснение";

            Excel.Range columnHeaderRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][1]];
            columnHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            columnHeaderRange.Font.Bold = true;
            startRowIndex++;

            for (int i = 0; i < result.Count(); i++)
            {
                string line = result[i];
                string col1 = "";
                string col2 = "";
                string col3 = "";

                int eqIndex = line.IndexOf('=');
                if (eqIndex != -1)
                {
                    col1 = line.Substring(0, eqIndex).TrimEnd();
                    string restAfterEq = line.Substring(eqIndex + 1).TrimStart();

                    int endIndex = restAfterEq.IndexOf("з");
                    if (endIndex != -1)
                    {
                        col2 = restAfterEq.Substring(0, endIndex).TrimEnd();
                        col3 = restAfterEq.Substring(endIndex).Trim();
                    }
                    else { col2 = restAfterEq; }
                }
                else { col1 = line; }

                worksheet.Cells[1][startRowIndex] = col1;
                worksheet.Cells[2][startRowIndex] = col2;
                worksheet.Cells[3][startRowIndex] = col3;

                startRowIndex++;
            }
            worksheet.Columns.AutoFit();
            application.Visible = true;
        }
    }
}

using System;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using OfficeOpenXml;
using System.Data.Entity;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.IO;

namespace Template_4435
{
    /// <summary>
    /// Логика взаимодействия для _4335Uteev.xaml
    /// </summary>
    public partial class _4335Uteev : Window
    {
        public _4335Uteev()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Создание объекта для чтения xlsx-файла
            var package = new ExcelPackage(new FileInfo("C:\\Учеба\\ИСРПО\\Импорт\\2.xlsx"));

            // Открытие соединения с базой данных
            using (airtimdelal connection = new airtimdelal())
            {
                // Обход строк в xlsx-файле
                var worksheet = package.Workbook.Worksheets[0];
                for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                {
                    // Чтение данных из xlsx-файла
                    string id = (worksheet.Cells[row, 1].Value.ToString());
                    string kodZakaza = (worksheet.Cells[row, 2].Value.ToString());
                    string dataSozdaniya = worksheet.Cells[row, 3].Value.ToString();
                    string kodKlienta = (worksheet.Cells[row, 4].Value.ToString());
                    string uslugi = worksheet.Cells[row, 5].Value.ToString();

                   

                    // Создание объекта для выполнения SQL-запроса
                    zakazy заказы = new zakazy(id,kodZakaza, dataSozdaniya, kodKlienta, uslugi);
                    connection.zakazies.Add(заказы);
                    connection.SaveChanges();
                }

            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            using (var connection = new SqlConnection("data source=DESKTOP-KB5JPL3;initial catalog=Orders;integrated security=True;MultipleActiveResultSets=True"))
            {
                connection.Open();

                // Выборка данных из базы данных
                var command = new SqlCommand("SELECT * FROM [zakazy]", connection);
                var dataReader = command.ExecuteReader();
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Создание файла Excel и заполнение его данными
                var excelPackage = new ExcelPackage();
                var worksheet = excelPackage.Workbook.Worksheets.Add("Worksheet Name");
                worksheet.Cells.LoadFromDataReader(dataReader, true);

                // Сохранение файла Excel на диск
                var file = new FileInfo("C:\\Users\\jonch\\OneDrive\\Рабочий стол\\ss.xlsx");
                excelPackage.SaveAs(file);

                MessageBox.Show("Данные экспортированы успешно!");
            }
        }
    }
}
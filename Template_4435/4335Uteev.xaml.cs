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
using Newtonsoft.Json;
using System.Text.Json;


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

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            using (var fileStream = new FileStream("C:\\Учеба\\ИСРПО\\Импорт\\2.json", FileMode.Open))
            {
                var options = new JsonSerializerOptions();
                serDeser[] data = System.Text.Json.JsonSerializer.Deserialize<serDeser[]>(fileStream);

                using (airtimdelal connection = new airtimdelal())
                {
                    foreach(var i in data)
                    {
                        connection.zakazies.Add(new zakazy(i.Id.ToString(), i.CodeOrder, i.CreateDate.ToString(), i.CodeClient, i.Services.ToString()));
                        connection.SaveChanges();
                    }
               
                }
            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            DataTable table = GetDataTableFromDatabase();

            // Создаем новый документ Word
            using (WordprocessingDocument doc = WordprocessingDocument.Create("example.docx", WordprocessingDocumentType.Document))
            {
                // Добавляем секцию
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Добавляем абзацы
                Paragraph para = body.AppendChild(new Paragraph());
                para.AppendChild(new Run(new Text("Данные из базы данных:")));

                // Добавляем таблицу с данными
                Table tableElement = new Table();
                TableRow headerRow = new TableRow();
                foreach (DataColumn column in table.Columns)
                {
                    headerRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(column.ColumnName)))));
                }
                tableElement.AppendChild(headerRow);
                foreach (DataRow row in table.Rows)
                {
                    TableRow tableRow = new TableRow();
                    foreach (DataColumn column in table.Columns)
                    {
                        tableRow.AppendChild(new TableCell(new Paragraph(new Run(new Text(row[column].ToString())))));
                    }
                    tableElement.AppendChild(tableRow);
                }
                body.AppendChild(tableElement);

                // Сохраняем документ
                doc.Save();
            }
        }
    }
}
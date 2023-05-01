using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.Json;
using Word = Microsoft.Office.Interop.Word;

namespace Template_4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_Savelev.xaml
    /// </summary>
    public partial class _4333_Savelev : Window
    {
        public _4333_Savelev()
        {
            InitializeComponent();
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл"
            };

            if (!(ofd.ShowDialog() == true))
                return;

            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            int _rows = ObjWorkSheet.Cells[ObjWorkSheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
            int _columns = ObjWorkSheet.Cells[1, ObjWorkSheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            list = new string[_rows, _columns];

            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 2, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (ISRPOEntities isrpoEntities = new ISRPOEntities())
            {
                for (int i = 0; i < _rows - 1; i++)
                {
                    DateTime dateOfCreate = DateTime.Parse(list[i, 2]);
                    TimeSpan timeShow = TimeSpan.Parse(list[i, 3]);
                    DateTime dateOfClose = new DateTime();
                    if (list[i, 7] != "")
                        dateOfClose = DateTime.Parse(list[i, 7]);
                    else
                        dateOfClose = Convert.ToDateTime(null);
                    isrpoEntities.Clients.Add(new Clients()
                    {
                        ID = Convert.ToInt32(list[i, 0]),
                        Код_Заказа = list[i, 1],
                        Дата_создания = dateOfCreate,
                        Время_показа = timeShow,
                        Код_Клиента = Convert.ToInt32(list[i, 4]),
                        Услуги = list[i, 5],
                        Статус = list[i, 6],
                        Дата_закрытия = dateOfClose,
                        Время_проката = list[i, 8],

                    });
                }
                isrpoEntities.SaveChanges();
                MessageBox.Show("Успешный импорт");
            }
        }

        private void Export_Click(object sender, RoutedEventArgs e)
        {
            List<Clients> category_1;
            List<Clients> category_2;
            List<Clients> category_3;
            List<Clients> category_4;
            List<Clients> category_5;
            List<Clients> category_6;
            List<Clients> category_7;


            using (ISRPOEntities isrpoEntities = new ISRPOEntities())
            {

                category_1 = isrpoEntities.Clients.Where(x => x.Время_проката == "2 часа" || x.Время_проката == "120 минут").ToList();
                category_2 = isrpoEntities.Clients.Where(x => x.Время_проката == "4 часа").ToList();
                category_3 = isrpoEntities.Clients.Where(x => x.Время_проката == "6 часов").ToList();
                category_4 = isrpoEntities.Clients.Where(x => x.Время_проката == "320 минут").ToList();
                category_5 = isrpoEntities.Clients.Where(x => x.Время_проката == "480 минут").ToList();
                category_6 = isrpoEntities.Clients.Where(x => x.Время_проката == "10 часов" || x.Время_проката == "600 минут").ToList();
                category_7 = isrpoEntities.Clients.Where(x => x.Время_проката == "12 часов").ToList();
            }

            var allCategories = new List<List<Clients>>()
            {
                category_1,
                category_2,
                category_3,
                category_4,
                category_5,
                category_6,
                category_7
            };

            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 7;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < 7; i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Категория {i + 1}";
                worksheet.Cells[1][startRowIndex] = "ID";
                worksheet.Cells[1][startRowIndex].Font.Bold = true;
                worksheet.Cells[2][startRowIndex] = "Код заказа";
                worksheet.Cells[2][startRowIndex].Font.Bold = true;
                worksheet.Cells[3][startRowIndex] = "Дата создания";
                worksheet.Cells[3][startRowIndex].Font.Bold = true;
                worksheet.Cells[4][startRowIndex] = "Код клиента";
                worksheet.Cells[4][startRowIndex].Font.Bold = true;
                worksheet.Cells[5][startRowIndex] = "Услуги";
                worksheet.Cells[5][startRowIndex].Font.Bold = true;

                foreach (var person in allCategories[i])
                {
                    startRowIndex++;
                    worksheet.Cells[1][startRowIndex] = person.ID;
                    worksheet.Cells[2][startRowIndex] = person.Код_Заказа;
                    worksheet.Cells[3][startRowIndex] = person.Дата_создания;
                    worksheet.Cells[4][startRowIndex] = person.Код_Клиента;
                    worksheet.Cells[5][startRowIndex] = person.Услуги;
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle =
                    Excel.XlLineStyle.xlContinuous;

                worksheet.Columns.AutoFit();
            }

            app.Visible = true;

        }


        class Person
        {
            public int Id { get; set; }
            public string CodeOrder { get; set; }
            public string CreateDate { get; set; }
            public string CreateTime { get; set; }
            public string CodeClient { get; set; }
            public string Services { get; set; }
            public string Status { get; set; }
            public string ClosedDate { get; set; }
            public string ProkatTime { get; set; }

        }

        private async void ImportJSON_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.json",
                Filter = "файл Json |*.json",
                Title = "Выберите файл"
            };

            if (!(ofd.ShowDialog() == true))
                return;

            List<Person> list;

            using (FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate))
            {
                list = await JsonSerializer.DeserializeAsync<List<Person>>(fs);
            }
            using (ISRPOEntities db = new ISRPOEntities())
            {

                foreach (Person person in list)
                {
                    DateTime DateCreate = DateTime.Parse(person.CreateDate.ToString());
                    TimeSpan TimeCreate = TimeSpan.Parse(person.CreateTime.ToString());
                    DateTime DateClosed =  new DateTime();

                    if (person.ClosedDate != "")
                        DateClosed = DateTime.Parse(person.ClosedDate.ToString());
                    else
                        DateClosed = Convert.ToDateTime(null);

                    db.Clients.Add(new Clients()
                    {
                        ID = person.Id,
                        Код_Заказа = person.CodeOrder,
                        Дата_создания = DateCreate,
                        Время_показа = TimeCreate,
                        Код_Клиента = Convert.ToInt32(person.CodeClient),
                        Услуги = person.Services,
                        Статус = person.Status,
                        Дата_закрытия = DateClosed,
                        Время_проката = person.ProkatTime
                    });

                }
                db.SaveChanges();
                MessageBox.Show("Успешный импорт");
            }
        }

        private void ExportWORD_Click(object sender, RoutedEventArgs e)
        {

            List<Clients> clients = new List<Clients>();
            using (ISRPOEntities db = new ISRPOEntities())
            {
                clients = db.Clients.ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();

                List<Clients> category_1;
                List<Clients> category_2;
                List<Clients> category_3;
                List<Clients> category_4;
                List<Clients> category_5;
                List<Clients> category_6;
                List<Clients> category_7;


                using (ISRPOEntities isrpoEntities = new ISRPOEntities())
                {
                    category_1 = isrpoEntities.Clients.Where(x => x.Время_проката == "2 часа" || x.Время_проката == "120 минут").ToList();
                    category_2 = isrpoEntities.Clients.Where(x => x.Время_проката == "4 часа").ToList();
                    category_3 = isrpoEntities.Clients.Where(x => x.Время_проката == "6 часов").ToList();
                    category_4 = isrpoEntities.Clients.Where(x => x.Время_проката == "320 минут").ToList();
                    category_5 = isrpoEntities.Clients.Where(x => x.Время_проката == "480 минут").ToList();
                    category_6 = isrpoEntities.Clients.Where(x => x.Время_проката == "10 часов" || x.Время_проката == "600 минут").ToList();
                    category_7 = isrpoEntities.Clients.Where(x => x.Время_проката == "12 часов").ToList();
                }

                var allCategories = new List<List<Clients>>()
                {
                     category_1,
                     category_2,
                     category_3,
                     category_4,
                     category_5,
                     category_6,
                     category_7
                };
                int i = 1;
                foreach (var category in allCategories)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = "Категория " + i;
                    i++;
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table clientTable = document.Tables.Add(tableRange, category.Count() + 1, 5);
                    clientTable.Borders.InsideLineStyle = clientTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    clientTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = clientTable.Cell(1, 1).Range;
                    cellRange.Text = "ID";
                    cellRange = clientTable.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = clientTable.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = clientTable.Cell(1, 4).Range;
                    cellRange.Text = "Код Клиента";
                    cellRange = clientTable.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    clientTable.Rows[1].Range.Bold = 1;
                    clientTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    int j = 1;
                    foreach (var person in category)
                    {
                        cellRange = clientTable.Cell(j + 1, 1).Range;
                        cellRange.Text = person.ID.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = clientTable.Cell(j + 1, 2).Range;
                        cellRange.Text = person.Код_Заказа.ToString(); ;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = clientTable.Cell(j + 1, 3).Range;
                        cellRange.Text = person.Дата_создания.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = clientTable.Cell(j + 1, 4).Range;
                        cellRange.Text = person.Код_Клиента.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = clientTable.Cell(j + 1, 5).Range;
                        cellRange.Text = person.Услуги.ToString();
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        j++;
                    }
                
                        Word.Paragraph DateParagraph = document.Paragraphs.Add();
                        Word.Range FirstDate = DateParagraph.Range;
                        Word.Range LastDate = DateParagraph.Range;
                        LastDate.Text = $"Дата последнего заказа - {category.Last().Дата_создания}";
                        LastDate.InsertParagraphAfter(); 
                        FirstDate.Text = $"Дата первого заказа - {category.First().Дата_создания}";
                        FirstDate.InsertParagraphAfter();
                        

                        document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                     
                }
                app.Visible = true;
                document.SaveAs2(@"D:\outputFileWord.docx");
                document.SaveAs2(@"D:\outputFilePdf.pdf", Word.WdExportFormat.wdExportFormatPDF);
            }
        }
    }
}

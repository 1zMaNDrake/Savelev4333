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
                worksheet.Cells[2][startRowIndex] = "Код клиента";
                worksheet.Cells[2][startRowIndex].Font.Bold = true;
                worksheet.Cells[3][startRowIndex] = "Время проката";
                worksheet.Cells[3][startRowIndex].Font.Bold = true;

                foreach (var person in allCategories[i])
                {
                    startRowIndex++;
                    worksheet.Cells[1][startRowIndex] = person.ID;
                    worksheet.Cells[2][startRowIndex] = person.Код_Клиента;
                    worksheet.Cells[3][startRowIndex] = person.Время_проката;
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[3][startRowIndex]];
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
    }
}

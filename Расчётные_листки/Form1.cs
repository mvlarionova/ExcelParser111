using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using MetroFramework.Forms;




namespace Расчётные_листки
{
    public partial class Form1 : MetroForm
    {
        //private 
        string[,] list = new string[31, 31]; // массив значений с листа равен по размеру листу


        public Form1()
        {
            InitializeComponent();
        }

        Task ProcessImport(List<string> data, IProgress<ProgressBar> progress)
        {
            int index = 1;
            int totalProgress = data.Count;
            var progressBar = new ProgressBar();
            return Task.Run(() =>
            {
                for (int i = 0; i < totalProgress; i++)
                {
                    progressBar.PercentComlete = index++ * 100 / totalProgress;
                    progress.Report(progressBar);
                    Thread.Sleep(15);
                }
            });

        }

        private async void metroButton1_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();

            for (int i = 0; i < 1000; i++)
                list.Add(i.ToString());
            metroLabel1.Text = "Working...";
            var progressBar = new Progress<ProgressBar>();

            progressBar.ProgressChanged += (o, report) =>
            {
                metroLabel1.Text = string.Format("Processing...{0}%",report.PercentComlete);
                metroProgressBar1.Value = report.PercentComlete;
                metroProgressBar1.Update();
            };
            await ProcessImport(list, progressBar);
            metroLabel1.Text = "Done!!!";
        }



        private async void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string name;
            if (metroComboBox1.SelectedItem.ToString() == "Выбрать другой путь")
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    name = openFileDialog1.FileName;
                    metroTextBox1.Clear();
                    metroTextBox1.Text = name;
                }
            }
            if (metroComboBox1.SelectedIndex == 0)
            {
                    metroTextBox1.Clear();
                metroTextBox1.Text = metroComboBox1.SelectedItem.ToString();
            }

            //List<string> list = new List<string>();
            

        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        private void ExportToExcel()
        {
            Excel.Application exApp = new Excel.Application(); //Откр документ

            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"E:\041.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            //-------------------------------------
            int lastColumn = (int)lastCell.Column;//!сохраним непосредственно требующееся в дальнейшем
            int lastRow = (int)lastCell.Row;
            //-------------------------------------
            for (int i = 0; i < 31; i++) //по всем колонкам
                for (int j = 0; j < 31; j++) // по всем строкам
                    list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку




            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой -- в том числе не используемые явно объекты !
            //for (int i = 1; i < lastColumn; i++) //по всем колонкам
            //    for (int j = 1; j < lastRow; j++) // по всем строкам 
            //        Console.Write(list[i, j]);//выводим строку
            //Console.ReadLine();
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application(); //Откр документ

            exApp.Workbooks.Add(); //Созд книгу
            exApp.Visible = false; // Видимое окно
            Excel.Worksheet workSheet = (Excel.Worksheet)exApp.ActiveSheet;
            Excel.Range range = workSheet.Rows; //Размеры ячеек, шрифт
            range.BorderAround2();
            range.Font.Name = "Times New Roman";
            range.Font.Bold = false;
            range.Font.Size = 14;

            int rowExcel2 = 1;


            range.Borders.LineStyle = true; //Стиль линий для изм

            workSheet.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlNoSelection;

            for (int i = 0; i < 31; i++)
            {
                for (int j = 0; j < 31; j++)
                {
                    workSheet.Cells[rowExcel2, "A"] = list[i,j];
                }
            }
            exApp.Quit(); //Закр док

            GC.Collect();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(exApp);
            int a=0;
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            dataGridView1.RowCount = 35;
            dataGridView1.ColumnCount = 35;
            for (int i = 0; i < 30; i++)
            {
                for (int j = 0; j < 30; j++)
                {
                    dataGridView1.Rows[i].Cells[j].Value = list[i, j];
                }
            }
        }
    }
}

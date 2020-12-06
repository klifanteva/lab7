using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace Lab5MSOffice
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void launchMSWordButton_Click(object sender, EventArgs e)
        {
            Word.Application appWord = new Word.Application();
            Word.Document docWord = null;
            string source = @"D:\Lab5.docx";
            docWord = appWord.Documents.Open(source);
            docWord.Activate();
            appWord.Visible = true;
            /*object missing = Type.Missing;
            Word.Document doc = app.Documents.Add(ref missing, ref missing, ref missing, ref missing);*/
            Word.Paragraph para = docWord.Paragraphs.Add();
            docWord.Paragraphs[1].Range.Bold = 2;
            docWord.Paragraphs[1].Range.Text = this.textBox2.Text;
        }

        private void launchMSExcelButton_Click(object sender, EventArgs e)
        {
            
            Excel.Application appExcel = new Excel.Application();
            Excel.Workbook workbook1;//книга откуда копировать
            Excel.Workbook workbook2;//книга куда копировать
            Excel.Worksheet workSheet1;//лист который копировать
            string source1 = @"D:\Lab5_1.xlsx";
            workbook1 = appExcel.Workbooks.Open(source1);
            string source2 = @"D:\Lab5_2.xlsx";
            workbook2 = appExcel.Workbooks.Open(source2);
            workSheet1 = workbook1.Worksheets["Лист1"];
            workSheet1.Copy(After: workbook2.Worksheets[workbook2.Worksheets.Count]);
            MessageBox.Show("Лист '" + workSheet1.Name.ToString() + "' успешно скопирован", "Поиск", MessageBoxButtons.OK, MessageBoxIcon.Information);
            appExcel.Visible = true;
            workbook1.Save();
            workbook1.Close();
            workbook2.Save();
            workbook2.Close();
            appExcel.Quit();
            
        }
    }
}

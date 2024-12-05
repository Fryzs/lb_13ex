using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
namespace lb_13ex
{
    public partial class Form1 : Form
    {
        int rowSelextedIndex;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook wkbk = excelApp.Workbooks.Add();
            excelApp = new Excel.Application();
            Excel.Worksheet sheet = wkbk.Sheets[1];


            for (int col = 0; col < dataGridView1.ColumnCount; col++)
            {
                sheet.Cells[1, col + 1] = dataGridView1.Columns[col].HeaderText;
                Excel.Range headerCell = sheet.Cells[1, col + 1];
                headerCell.Font.Bold = true;
                headerCell.Interior.Color = Excel.XlRgbColor.rgbLightGray;
                headerCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }

            for (int row = 0; row < dataGridView1.RowCount; row++)
            {
                for (int col = 0; col < dataGridView1.ColumnCount; col++)
                {
                    if (dataGridView1.Rows[row].Cells[col].Value != null)
                        sheet.Cells[row + 2, col + 1] = dataGridView1.Rows[row].Cells[col].Value.ToString();

                    Excel.Range dataCell = sheet.Cells[row + 3, col + 1];
                    dataCell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//Обвод клітинки
                }
            }
            sheet.Columns.AutoFit(); //Авто size
        }

        private void button3_Click(object sender, EventArgs e)
        {
            int rowIndex = dataGridView1.Rows.Add();

            dataGridView1.Rows[rowIndex].Cells[0].Value = comboBox1.Text;
            dataGridView1.Rows[rowIndex].Cells[1].Value = textBox1.Text;
            dataGridView1.Rows[rowIndex].Cells[2].Value = textBox2.Text;
            dataGridView1.Rows[rowIndex].Cells[3].Value = textBox3.Text;


        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 1 && !dataGridView1.Rows[dataGridView1.Rows.Count - 1].IsNewRow)
                dataGridView1.Rows.RemoveAt(dataGridView1.RowCount - 1);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            rowSelextedIndex = e.RowIndex;

            comboBox1.Text = dataGridView1.Rows[rowSelextedIndex].Cells[0].Value.ToString();
            textBox1.Text = dataGridView1.Rows[rowSelextedIndex].Cells[1].Value.ToString();
            textBox2.Text = dataGridView1.Rows[rowSelextedIndex].Cells[2].Value.ToString();
            textBox3.Text = dataGridView1.Rows[rowSelextedIndex].Cells[3].Value.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows[rowSelextedIndex].Cells[0].Value = comboBox1.Text;
            dataGridView1.Rows[rowSelextedIndex].Cells[1].Value = textBox1.Text;
            dataGridView1.Rows[rowSelextedIndex].Cells[2].Value = textBox2.Text;
            dataGridView1.Rows[rowSelextedIndex].Cells[3].Value = textBox3.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(openFileDialog1.FileName);
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];

                int rowCount = sheet.UsedRange.Rows.Count;
                int colCount = sheet.UsedRange.Columns.Count;

                while (rowCount >= dataGridView1.RowCount)
                    dataGridView1.Rows.Add();

                for (int row = 2; row <= rowCount; row++)
                {

                    for (int col = 1; col <= colCount; col++)
                    {

                        var cellValue = sheet.Cells[row, col].Text;
                        dataGridView1.Rows[row - 2].Cells[col - 1].Value = cellValue;

                    }
                }
                workbook.Close(false);
                excelApp.Quit();
            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}

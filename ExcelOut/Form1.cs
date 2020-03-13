using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Exel = Microsoft.Office.Interop.Excel;

namespace ExcelOut
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView1.Columns.Add("colum1", "Столбец1");
            dataGridView1.Columns.Add("colum2", "Столбец2");
            dataGridView1.Columns.Add("colum3", "Столбец3");
        }

        private void buttoAdd_Click(object sender, EventArgs e)
        {

            dataGridView1.Rows.Add("1", "2", "3");
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            if (!(saveFileDialog.ShowDialog() == DialogResult.OK))
            {
                MessageBox.Show("error"); return;
            }
            string FilePath = saveFileDialog.FileName;

            Exel.Application application = new Exel.Application();

            Exel.Workbook workbook = application.Workbooks.Add();
            Exel.Worksheet worksheet;
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "лист1";

            for (int i = 1; i <= dataGridView1.RowCount; i++)
            {
                for (int j = 1; j <= dataGridView1.ColumnCount; j++)
                {
                    worksheet.Cells[i, j] = dataGridView1.Rows[i - 1].Cells[j - 1];
                }
            }

            application.Application.ActiveWorkbook.SaveAs(FilePath);
        }
    }
}

using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace ImportarPlanoContas
{
    public partial class ASCIItoExcel : Form
    {
        public ASCIItoExcel()
        {
            InitializeComponent();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog _openFileDialog = new OpenFileDialog())
            {
                if (_openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    toolStripTextBox1.Text = _openFileDialog.FileName;
                    GetASCIINames();
                }
            }
        }

        private void GetASCIINames()
        {
            string[] lines = File.ReadAllLines(toolStripTextBox1.Text);
            string[] values;

            DataTable _table = new DataTable();
            
            for (int i = 0; i < lines.Length; i++)
            {
                values = lines[i].ToString().Split(',');
                string[] row = new string[values.Length];

                for (int j = 0; j < values.Length; j++)
                {
                    if (i == 0)
                    {
                        _table.Columns.Add(values[j].Trim());
                    }

                    else
                    {
                        row[j] = values[j].Trim();

                        
                    }

                    
                }

                _table.Rows.Add(row);





                //dataGridView1.Rows.Add(row);
            }

            dataGridView1.DataSource = _table;
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            ExceltoASCII form = new ExceltoASCII();
            this.Hide();
            form.ShowDialog();
            this.Close();
        }

        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            copyAlltoClipboard();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            //if (dataGridView1.Rows.Count > 0)
            //{
            //    Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
            //    xcelApp.Application.Workbooks.Add(Type.Missing);

            //    int _iColCount = 1;

            //    foreach(DataGridViewColumn _col in dataGridView1.Columns)
            //    {
            //        xcelApp.Cells[1, _iColCount] = _col.HeaderText;
            //        _iColCount++;
            //    }

            //    //for (int i = 0; i < dataGridView1.Columns.Count; i++)
            //    //{
            //      //  xcelApp.Cells[1, i + 1] = dataGridView1.Columns[i - 1].HeaderText;
            //    //}

            //    for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //    {
            //        for (int j = 0; j < dataGridView1.Columns.Count; j++)
            //        {
            //            xcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
            //        }
            //    }
            //    xcelApp.Columns.AutoFit();
            //    xcelApp.Visible = true;
            //}
        }
    }
}

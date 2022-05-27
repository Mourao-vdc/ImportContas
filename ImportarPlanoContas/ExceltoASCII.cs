using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ImportarPlanoContas
{
    public partial class ExceltoASCII : Form
    {
        public string header { get; set; }
        public string Texto { get; set; }
        public ExceltoASCII()
        {
            InitializeComponent();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            txtExport.Text = "";
            toolStripComboBox1.Items.Clear();

            using (OpenFileDialog _openFileDialog = new OpenFileDialog())
            {
                if (_openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    toolStripTextBox1.Text = _openFileDialog.FileName;

                    if (toolStripComboBox1.Items.Count > 0)
                        toolStripComboBox1.ComboBox.SelectedIndex = 0;
                }
            }
        }

        private void toolStripTextBox1_TextChanged(object sender, EventArgs e)
        {
            string[] _sNomes = GetExcelSheetNames(toolStripTextBox1.Text);

            foreach (string _sNome in _sNomes)
            {
                toolStripComboBox1.Items.Add(_sNome);

            }
        }

        public String[] GetExcelSheetNames(string excelFile, string _sCon = "")
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            try
            {
                //Conexão com o Excel:
                String connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFile + ";Extended Properties=Excel 8.0;";

                if (!string.IsNullOrEmpty(_sCon))
                    connString = _sCon;

                //String connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + excelFile + "; Extended Properties = Excel 12.0;";

                //Cria um objeto de conexão:
                objConn = new OleDbConnection(connString);

                //Abre a conexão:
                objConn.Open();

                //Obter a tabela com os dados das schemas:
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                //Cria um array de strings:
                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                //Adiciona o nome da folha para um array:
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                return excelSheets;
            }

            catch (OleDbException)
            {
                try
                {
                    String connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + excelFile + "; Extended Properties = Excel 12.0;";

                    //Cria um objeto de conexão:
                    objConn = new OleDbConnection(connString);

                    //Abre a conexão:
                    objConn.Open();

                    //Obter a tabela com os dados das schemas:
                    dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    if (dt == null)
                    {
                        return null;
                    }

                    //Cria um array de strings:
                    String[] excelSheets = new String[dt.Rows.Count];
                    int i = 0;

                    //Adiciona o nome da folha para um array:
                    foreach (DataRow row in dt.Rows)
                    {
                        excelSheets[i] = row["TABLE_NAME"].ToString();
                        i++;
                    }

                    return excelSheets;
                }

                catch
                {
                    return null;
                }
            }

            catch
            {
                return null;
            }

            finally
            {
                //Fecha e limpa as conexões:
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //Cria a string de conexão:
                string _sPath = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + toolStripTextBox1.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";

                //Cria a conexão:
                using (OleDbConnection _con = new OleDbConnection(_sPath))
                {
                    //Obtém os dados da tabela escolhida:
                    using (OleDbDataAdapter _dataAdapter = new OleDbDataAdapter("SELECT * FROM [" + toolStripComboBox1.SelectedItem.ToString() + "]", _con))
                    {
                        //Cria a dt:
                        using (System.Data.DataTable _dt = new System.Data.DataTable())
                        {
                            //Preenche a dt:
                            _dataAdapter.Fill(_dt);

                            //Associa a dt à GDV:
                            dataGridView1.DataSource = _dt;

                            _con.Close();
                        }
                    }
                }
            }

            catch
            { }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            this.Size = new Size(1470, 489);
            
            foreach(DataGridViewColumn _col in dataGridView1.Columns)
            {
                header += _col.HeaderText + ",";
                //MessageBox.Show(header);
            }

            if (header.EndsWith(","))
                header = header.Substring(0, header.Length - 1).Replace("#",".");

            string _sAscii = header + Environment.NewLine;

            foreach (DataGridViewRow _row in dataGridView1.Rows)
            {
                string row = string.Empty;

                foreach (DataGridViewColumn _col in dataGridView1.Columns)
                {
                    row += _row.Cells[_col.Index].Value.ToString().Replace(",",".") + ",";

                    //string _sNumConta = ;
                    //string _sTipoConta = _row.Cells[1].Value.ToString();
                    //string _sNome = _row.Cells[2].Value.ToString();
                    //string _sTaxonomia = _row.Cells[3].Value.ToString();

                    //string _sTaxonomiaFinal = "0";

                    /*try
                    {
                        if (!string.IsNullOrEmpty(_sNumConta))
                        {
                            string[] _sSplitedTaxonomia = _sNumConta.Split('-');

                            _sTaxonomiaFinal = _sSplitedTaxonomia[0].Trim();
                        }
                    }

                    catch
                    {
                        _sTaxonomiaFinal = "0";
                    }*/


                    //txtExport.Text = _sAscii + _sNumConta + ", " + _sTaxonomiaFinal + Environment.NewLine;
                }

                if(row.EndsWith(","))
                    row = row.Substring(0, row.Length - 1);

                _sAscii += row + Environment.NewLine;
            }

            txtExport.Text = _sAscii.Remove(_sAscii.Length - 1);

            Clipboard.SetText(_sAscii);

            //MessageBox.Show("Test");
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            StreamWriter sw;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.FileName = "Novo Documento";
            saveFileDialog1.Filter = "txt files (*.ASC)|*.ASC|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 1;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if(File.Exists(saveFileDialog1.FileName + saveFileDialog1.Filter))
                {
                    MessageBox.Show("Já existe um documento com esse nome, deseja substitui-lo?","Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                    if(DialogResult == DialogResult.Yes)
                    {
                        sw = File.CreateText(saveFileDialog1.FileName);
                        string text = txtExport.Text;
                        sw.WriteLine(text);
                        sw.Close();
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    sw = File.CreateText(saveFileDialog1.FileName);
                    string text = txtExport.Text;
                    sw.WriteLine(text);
                    sw.Close();
                }
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            ASCIItoExcel form = new ASCIItoExcel();
            this.Hide();
            form.ShowDialog();
            this.Close();
        }
    }
}
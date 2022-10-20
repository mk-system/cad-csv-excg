using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace CadCsvExcg
{
    public partial class frmPreview : Form
    {
        public frmPreview()
        {
            InitializeComponent();
        }

        public void loadCSV(string file, string delimiter, bool header, int id, bool append, string colname = "COLUMN")
        {
            try
            {
                this.Text = file;
                DataTable dt = CSVUtlity.Parse(file, delimiter, header, id, append, colname);
                this.dataGridView1.DataSource = dt;
                this.toolStripStatusLabel1.Text = "Columns: " + dt.Columns.Count;
                this.toolStripStatusLabel2.Text = "Rows: " + dt.Rows.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        public void loadDataTable(DataTable dt)
        {
            this.dataGridView1.DataSource = dt;
            this.toolStripStatusLabel1.Text = "Columns: " + dt.Columns.Count;
            this.toolStripStatusLabel2.Text = "Rows: " + dt.Rows.Count;
        }

        private void frmPreview_FormClosing(object sender, FormClosingEventArgs e)
        {
            dataGridView1.DataSource = null;
            this.Text = "Preview";
        }

        private void frmPreview_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            bool useHeader = true;
            DialogResult dialogResult = MessageBox.Show("Would you like to include header row in saving file?", "Include header", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                useHeader = true;
            }
            else if (dialogResult == DialogResult.No)
            {
                useHeader = false;
            }
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "CSV Files (*.csv)|*.csv";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                ToCSV((DataTable)dataGridView1.DataSource, saveFileDialog1.FileName, useHeader);
            }
            
        }

        private void ToCSV(DataTable dtDataTable, string strFilePath, bool header = true)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers
            if (header)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    sw.Write(dtDataTable.Columns[i]);
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(","))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
    }
}

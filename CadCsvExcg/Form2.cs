using System;
using System.Data;
using System.Windows.Forms;

namespace CadCsvExcg
{
    public partial class frmPreview : Form
    {
        public frmPreview()
        {
            InitializeComponent();
        }

        public void loadCSV(string file, string delimiter, bool header, int id, bool append)
        {
            try
            {
                this.Text = file;
                DataTable dt = Csv.Parse(file, delimiter, header, id, append);
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
    }
}

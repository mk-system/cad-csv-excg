using System;
using System.Data;
using System.IO;
using System.Windows.Forms;

namespace CadCsvExcg
{
    public partial class FormPreview : Form
    {
        public FormPreview(DataTable dataTable)
        {
            InitializeComponent();
            if (dataTable != null)
            {
                this.dataGridView1.DataSource = dataTable;
                this.toolStripStatusLabel1.Text = "Columns: " + dataTable.Columns.Count;
                this.toolStripStatusLabel2.Text = "Rows: " + dataTable.Rows.Count;
            }
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

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CadCsvExcg
{
    public partial class frmPreview : Form
    {
        public string path;
        private DataTable table;
        public frmPreview()
        {
            InitializeComponent();
        }

        private void frmPreview_FormClosing(object sender, FormClosingEventArgs e)
        {
            dataGridView1.DataSource = null;
            path = null;
            table = null;
        }

        private void frmPreview_Shown(object sender, EventArgs e)
        {
            this.dataGridView1.Visible = false;
            this.backgroundWorker1.RunWorkerAsync();
        }


        private void backgroundWorker1_ProgressChanged_1(object sender, ProgressChangedEventArgs e)
        {
   
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ReadCSV csv = new ReadCSV(this.path);
            try
            {
                this.table = csv.readCSV;
            } catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.dataGridView1.DataSource = this.table;
            this.dataGridView1.AutoResizeColumns();
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.Visible = true;
        }
    }
}

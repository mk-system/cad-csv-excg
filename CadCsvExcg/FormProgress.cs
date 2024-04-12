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
    public partial class FormProgress : Form
    {

        public event EventHandler CancelEvent;

        public FormProgress()
        {
            InitializeComponent();
        }

        public void UpdateProgressBar(int value)
        {
            value = Math.Max(0, Math.Min(value, progressBar1.Maximum));
            progressBar1.Value = value;
            this.Text = value.ToString() + "%";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            CancelEvent?.Invoke(this, EventArgs.Empty);
        }
    }
}

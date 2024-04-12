using System;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using CsvHelper;
using CsvHelper.Configuration;

namespace CadCsvExcg
{
    public partial class FormPreview : Form, IDisposable
    {
        private string _filePath;
        
        private Delimiter _delimiter;
        
        private bool _hasHeader;
        
        private BackgroundWorker _worker;
        
        private BindingSource _bindingSource;

        private DataGridView _dataGridView;

        public FormPreview(
            string filePath,
            Delimiter delimiter = Delimiter.COMMA,
            bool hasHeader = false
            )
        {
            InitializeComponent();
            _filePath = filePath;
            _delimiter = delimiter;
            _hasHeader = hasHeader;

            this.Controls.Clear();

            _worker = new BackgroundWorker();
            _worker.WorkerReportsProgress = true;
            _worker.DoWork += ReadCsvFile;
            _worker.ProgressChanged += ReportProgress;
            _worker.RunWorkerCompleted += WorkCompleted;

            _bindingSource = new BindingSource();

            _dataGridView = new DataGridView();
            _dataGridView.ColumnHeadersVisible = true;
            _dataGridView.Dock = DockStyle.Fill;
            _dataGridView.ReadOnly = true;
            _dataGridView.AllowUserToAddRows = false;
            _dataGridView.AllowUserToDeleteRows = false;
            _dataGridView.DataSource = _bindingSource;
        }

        public void Start()
        {
            if (_worker.IsBusy != true)
            {
                this.Text = "Loading";
                _worker.RunWorkerAsync();
            }
        }

        private void ReadCsvFile(object sender, DoWorkEventArgs e)
        {
            try
            {
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = _hasHeader,
                    Delimiter = _delimiter.GetString(),
                    MissingFieldFound = null,
                };

                using (var stream = new MemoryStream())
                using (var reader = new StreamReader(_filePath))
                using (var csv = new CsvReader(reader, config))
                {
                    using (var dr = new CsvDataReader(csv))
                    {
                        var dataTable = new DataTable();
                        dataTable.Load(dr);
                        _bindingSource.DataSource = dataTable;
                    }
                }
                e.Result = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private void ReportProgress(object sender, ProgressChangedEventArgs e)
        {
            Console.WriteLine($"Progress: {e.ProgressPercentage}%");
        }

        private void WorkCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                // Handle the error
                MessageBox.Show($"An error occurred: {e.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return;
            }
            else if (e.Cancelled)
            {
                // Handle cancellation
                MessageBox.Show("Cancelled");
                this.Close();
                return;
            } else if ((bool)e.Result == true)
            {
                this.Controls.Add(_dataGridView);
                this.Text = _filePath;
            }
        }

        private void FormPreview2_Load(object sender, EventArgs e)
        {
            this.Start();
        }

        private void FormPreview2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_worker.IsBusy)
            {
                _worker.CancelAsync();
            }
            _bindingSource.Dispose();
            _dataGridView.Dispose();
            this.Controls.Clear();
        }
    }
}

using System.Data.SQLite;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Linq;

namespace CadCsvExcg
{
    public partial class frmMain : Form
    {
        private class ComboItem
        {
            public Delimiter ID { get; set; }
            public string Text { get; set; }
        }
        public frmMain()
        {
            InitializeComponent();
            // InitializeDatabase();

            var delimiters = new ComboItem[] {
                new ComboItem{ ID = Delimiter.COMMA, Text = "Comma (,)" },
                new ComboItem{ ID = Delimiter.TAB, Text = "Tab (    )" },
            };
            cobDelimiter1.DataSource = delimiters.Clone();
            cobDelimiter2.DataSource = delimiters.Clone();
            cobDelimiter1.DisplayMember = "Text";
            cobDelimiter1.ValueMember = "ID";
            cobDelimiter2.DisplayMember = "Text";
            cobDelimiter2.ValueMember = "ID";
        }


        static void InitializeDatabase()
        {
            using (var connection = new SQLiteConnection("Data Source=database.db;New=True;Compress=True;"))
            {
                try
                {
                    connection.Open();
                    var command = new SQLiteCommand(connection);

                    // Clearing exists tables
                    command.CommandText = "DROP TABLE IF EXISTS table1";
                    command.ExecuteNonQuery();
                    command.CommandText = "DROP TABLE IF EXISTS table2";
                    command.ExecuteNonQuery();

                    // Creating tables for storing data
                    command.CommandText = @"CREATE TABLE table1(id INTEGER PRIMARY KEY)";
                    command.ExecuteNonQuery();
                    command.CommandText = @"CREATE TABLE table2(id INTEGER PRIMARY KEY)";
                    command.ExecuteNonQuery();

                    SQLiteDataAdapter da = new SQLiteDataAdapter(command);
                    // da.Fill(dataTable);

                    connection.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (CloseCancel() == false)
            {
                e.Cancel = true;
            }
        }

        private static bool CloseCancel()
        {
            const string message = "Are you sure that you would like to close the program?";
            const string caption = "Closing the program";
            var result = MessageBox.Show(message, caption,
                                         MessageBoxButtons.YesNo,
                                         MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                return true;
            else
                return false;
        }

        private void btnExit_Click(object sender, EventArgs e) => System.Windows.Forms.Application.Exit();

        private void btnAdd1_Click(object sender, EventArgs e)
        {
            AddFile(lvCad);
        }

        private void btnAdd2_Click(object sender, EventArgs e)
        {
            AddFile(lvBom);
        }

        private void AddFile(ListView listView)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "CSV Files (*.csv)|*.csv";
            dialog.Multiselect = true;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                foreach (var fileName in dialog.FileNames)
                {
                    try
                    {
                        if (listView.FindItemWithText(fileName) == null)
                        {
                            ListViewItem item;
                            string[] arr = new string[3];
                            arr[0] = System.IO.Path.GetFileName(fileName);
                            arr[1] = fileName;
                            arr[2] = (new FileInfo(fileName).Length / 1024).ToString() + " KB";
                            item = new ListViewItem(arr);
                            listView.Items.Add(item);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                ResizeListViewColumnsWidth();
            }
        }

        private void RemoveSelectedItems(ListView listView)
        {
            foreach (ListViewItem item in listView.SelectedItems)
            {
                listView.Items.Remove(item);
            }
        }

        private void ResizeListViewColumnsWidth()
        {
            lvCad.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            lvCad.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
            lvBom.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            lvBom.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        private void btnRemove1_Click(object sender, EventArgs e)
        {
            RemoveSelectedItems(lvCad);
        }

        private void btnRemove2_Click(object sender, EventArgs e)
        {
            RemoveSelectedItems(lvBom);
        }

        private void lvCad_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnRemove1.Enabled = !!(lvCad.SelectedItems.Count > 0);
            btnPreview1.Enabled = !!(lvCad.SelectedItems.Count == 1);
        }

        private void lvBom_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnRemove2.Enabled = !!(lvBom.SelectedItems.Count > 0);
            btnPreview2.Enabled = !!(lvBom.SelectedItems.Count == 1);
        }

        private void btnSelectDir_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                llblOutputDir.Text = dialog.SelectedPath;
            }
        }

        private void llblOutputDir_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                var fullPath = System.IO.Path.GetFullPath(llblOutputDir.Text);
                Process.Start("explorer.exe", @"" + fullPath + "");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void frmMain_ResizeEnd(object sender, EventArgs e)
        {
            ResizeListViewColumnsWidth();
        }

        private void frmMain_MaximizedBoundsChanged(object sender, EventArgs e)
        {
            ResizeListViewColumnsWidth();
        }

        private void btnPreview1_Click(object sender, EventArgs e)
        {

            if (lvCad.SelectedItems.Count == 1)
            {
                Preview(lvCad.SelectedItems[0].SubItems[1].Text, (Delimiter)cobDelimiter1.SelectedValue, cbHeader1.Checked, (int)numCad.Value, cbUseFileName1.Checked, "A COLUMN");
            }
        }

        private void btnPreview2_Click(object sender, EventArgs e)
        {
            if (lvBom.SelectedItems.Count == 1)
            {
                Preview(lvBom.SelectedItems[0].SubItems[1].Text, (Delimiter)cobDelimiter2.SelectedValue, cbHeader2.Checked, (int)numBom.Value, cbUseFileName2.Checked, "B COLUMN");
            }
        }

        private void lvCad_DoubleClick(object sender, EventArgs e)
        {
            if (lvCad.SelectedItems.Count == 1)
            {
                Preview(lvCad.SelectedItems[0].SubItems[1].Text, (Delimiter)cobDelimiter1.SelectedValue, cbHeader1.Checked, (int)numCad.Value, cbUseFileName1.Checked, "A COLUMN");
            }
        }

        private void lvBom_DoubleClick(object sender, EventArgs e)
        {
            if (lvBom.SelectedItems.Count == 1)
            {
                Preview(lvBom.SelectedItems[0].SubItems[1].Text, (Delimiter)cobDelimiter2.SelectedValue, cbHeader2.Checked, (int)numBom.Value, cbUseFileName2.Checked, "B COLUMN");
            }
        }
        private void Preview(string path, Delimiter delimiter, bool header, int id = 0, bool append = false, string colname = "COLUMN")
        {
            // Csv.Parse(file, delimiter1, header1, (int)numCad.Value, cbUseFileName.Checked)
            if (File.Exists(path))
            {
                using (frmPreview frm = new frmPreview()) {
                    frm.loadCSV(path, delimiter.GetString(), header, id, append, colname);
                    frm.ShowDialog();
                }
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt1 = new DataTable();
                DataTable dt2 = new DataTable();
                DataTable dtResult = new DataTable();
                string delimiter1 = ((Delimiter)cobDelimiter1.SelectedValue).GetString();
                string delimiter2 = ((Delimiter)cobDelimiter2.SelectedValue).GetString();
                bool header1 = cbHeader1.Checked;
                bool header2 = cbHeader2.Checked;
                // merging unidrafCAD CSV files into dataTable
                foreach (ListViewItem item in lvCad.Items)
                {
                    string file = item.SubItems[1].Text;
                    dt1.Merge(CSVUtlity.Parse(file, delimiter1, header1, (int)numCad.Value, cbUseFileName1.Checked, "A COLUMN"));
                }
                // merging visualBOM CSV files into dataTable
                foreach (ListViewItem item in lvBom.Items)
                {
                    string path = item.SubItems[1].Text;
                    dt2.Merge(CSVUtlity.Parse(path, delimiter2, header2, (int)numBom.Value, cbUseFileName2.Checked, "B COLUMN"));
                }

                // merging two tables via primary id
                dtResult.Merge(dt1);
                dtResult.Merge(dt2);

                // showing result
                using (frmPreview frm = new frmPreview())
                {
                    frm.loadDataTable(dtResult);
                    frm.ShowDialog();
                }
            } catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

using System.Data.SQLite;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Linq;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Threading;
using System.Collections.Generic;

namespace CadCsvExcg
{
    public partial class FormMain : Form
    {

        private bool initFinished = false;
        private class ComboItem
        {
            public int ID { get; set; }
            public string Text { get; set; }
        }
        public FormMain()
        {
            InitializeComponent();
            // InitializeDatabase();

            var delimiters = new ComboItem[] {
                new ComboItem{ ID = (int)Delimiter.COMMA, Text = "Comma (,)" },
                new ComboItem{ ID = (int)Delimiter.TAB, Text = "Tab (    )" },
            };

            var encoders = new ComboItem[]
            {
                new ComboItem{ ID = 0, Text = "UTF-7" },
                new ComboItem{ ID = 1, Text = "UTF-8" },
                new ComboItem{ ID = 2, Text = "UTF-16LE" },
                new ComboItem{ ID = 3, Text = "UTF-16BE" },
                new ComboItem{ ID = 4, Text = "UTF-32" },
            };

            cobDelimiter1.DataSource = delimiters.Clone();
            cobDelimiter1.DisplayMember = "Text";
            cobDelimiter1.ValueMember = "ID";
            cobDelimiter2.DataSource = delimiters.Clone();
            cobDelimiter2.DisplayMember = "Text";
            cobDelimiter2.ValueMember = "ID";
            cobDelimiter3.DataSource = delimiters.Clone();
            cobDelimiter3.DisplayMember = "Text";
            cobDelimiter3.ValueMember = "ID";
            cobEncoding.DataSource = encoders;
            cobEncoding.DisplayMember = "Text";
            cobEncoding.ValueMember = "ID";

            LoadConfig();

            this.txtColumn1.Enabled = this.cbUseFileName1.Checked;
            this.txtColumn2.Enabled = this.cbUseFileName2.Checked;

            CheckCombinable();

            initFinished = true;
        }

        private void LoadConfig()
        {
            this.cbHeader1.Checked = !!Properties.Settings.Default.header1;
            this.cbHeader2.Checked = !!Properties.Settings.Default.header2;
            this.cbUseFileName1.Checked = !!Properties.Settings.Default.filename1;
            this.cbUseFileName2.Checked = !!Properties.Settings.Default.filename2;
            this.cobDelimiter1.SelectedIndex = (int)Properties.Settings.Default.delimiter1;
            this.cobDelimiter2.SelectedIndex = (int)Properties.Settings.Default.delimiter2;
            this.cobDelimiter3.SelectedIndex = (int)Properties.Settings.Default.delimiter3;
            this.numCad.Value = (int)Properties.Settings.Default.id1;
            this.numBom.Value = (int)Properties.Settings.Default.id2;
            this.txtColumn1.Text = Properties.Settings.Default.column1;
            this.txtColumn2.Text = Properties.Settings.Default.column2;
            this.llblOutputDir.Text = Properties.Settings.Default.output;
            this.cobEncoding.SelectedIndex = Properties.Settings.Default.encoding;
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
            CheckCombinable();
        }

        private void RemoveSelectedItems(ListView listView)
        {
            foreach (ListViewItem item in listView.SelectedItems)
            {
                listView.Items.Remove(item);
            }
            CheckCombinable();
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

        private void CheckCombinable()
        {
            grpSettings.Enabled = lvCad.Items.Count > 0 && lvBom.Items.Count > 0;
            btnStart.Enabled = lvCad.Items.Count > 0 && lvBom.Items.Count > 0;
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
                saveConfig();
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

        private void btnPreview1_Click(object sender, EventArgs e)
        {
            if (lvCad.SelectedItems.Count == 1)
            {
                Preview(lvCad.SelectedItems[0].SubItems[1].Text, (Delimiter)cobDelimiter1.SelectedValue, 0, cbHeader1.Checked, (int)numCad.Value, cbUseFileName1.Checked, txtColumn1.Text);
            }
        }

        private void btnPreview2_Click(object sender, EventArgs e)
        {
            if (lvBom.SelectedItems.Count == 1)
            {
                Preview(lvBom.SelectedItems[0].SubItems[1].Text, (Delimiter)cobDelimiter2.SelectedValue, 0, cbHeader2.Checked, (int)numBom.Value, cbUseFileName2.Checked, txtColumn2.Text);
            }
        }

        private void lvCad_DoubleClick(object sender, EventArgs e)
        {
            if (lvCad.SelectedItems.Count == 1)
            {
                Preview(lvCad.SelectedItems[0].SubItems[1].Text, (Delimiter)cobDelimiter1.SelectedValue, 0, cbHeader1.Checked, (int)numCad.Value, cbUseFileName1.Checked, txtColumn1.Text);
            }
        }

        private void lvBom_DoubleClick(object sender, EventArgs e)
        {
            if (lvBom.SelectedItems.Count == 1)
            {
                Preview(lvBom.SelectedItems[0].SubItems[1].Text, (Delimiter)cobDelimiter2.SelectedValue, 0, cbHeader2.Checked, (int)numBom.Value, cbUseFileName2.Checked, txtColumn2.Text);
            }
        }
        private void Preview(string path, Delimiter delimiter, int columnNameStartNum, bool header, int id = 0, bool append = false, string fileColumnName = "FILENAME")
        {
            Loading(true);
            using (FormProcess frm = new FormProcess())
            {
                List<object> args = new List<object>
                {
                    path,
                    delimiter.GetString(),
                    header,
                    columnNameStartNum,
                    id,
                    append,
                    fileColumnName
                };
                frm.Preview(args);
                frm.ShowDialog();
            }
            Loading(false);
            /*  if (File.Exists(path))
              {
                  using (FormPreview frm = new FormPreview())
                  {
                      Loading(true);
                      frm.LoadCSV(path, delimiter.GetString(), columnNameStartNum, header, id, append, fileColumnName);
                      Loading(false);
                      frm.ShowDialog();
                  }
              }*/
        }


        private void btnStart_Click(object sender, EventArgs e)
        {
            using (FormProcess frm = new FormProcess())
            {
                List<string> paths1 = new List<string>();
                List<string> paths2 = new List<string>();
                string delimiter1 = ((Delimiter)cobDelimiter1.SelectedValue).GetString();
                string delimiter2 = ((Delimiter)cobDelimiter2.SelectedValue).GetString();
                bool header1 = cbHeader1.Checked;
                bool header2 = cbHeader2.Checked;
                int primary1 = (int)numCad.Value;
                int primary2 = (int)numBom.Value;
                bool extra1 = cbUseFileName1.Checked;
                bool extra2 = cbUseFileName2.Checked;
                string extraName1 = txtColumn1.Text;
                string extraName2 = txtColumn2.Text;

                foreach (ListViewItem item in lvCad.Items)
                {
                    string file = item.SubItems[1].Text;
                    paths1.Add(file);
                }

                foreach (ListViewItem item in lvBom.Items)
                {
                    string file = item.SubItems[1].Text;
                    paths2.Add(file);
                }

                List<object> args = new List<object>
                    {
                        paths1.ToArray<string>(),
                        paths2.ToArray<string>(),
                        delimiter1,
                        delimiter2,
                        header1,
                        header2,
                        primary1,
                        primary2,
                        extra1,
                        extra2,
                        extraName1,
                        extraName2
                    };

                frm.Combine(args);
                frm.ShowDialog();
            }
            try
            {



                /*DataTable dt1 = new DataTable();
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
                    dt1.Merge(CSVUtility.Import(file, delimiter1, 0, header1, (int)numCad.Value, cbUseFileName1.Checked, txtColumn1.Text));
                }
                // merging visualBOM CSV files into dataTable
                foreach (ListViewItem item in lvBom.Items)
                {
                    string path = item.SubItems[1].Text;
                    dt2.Merge(CSVUtility.Import(path, delimiter2, dt1.Columns.Count, header2, (int)numBom.Value, cbUseFileName2.Checked, txtColumn2.Text));
                }

                // merging two tables via primary id
                dtResult.Merge(dt1);
                dtResult.Merge(dt2);*/

                // showing result
                /*using (FormPreview frm = new FormPreview())
                {
                    frm.LoadDataTable(dtResult);
                    frm.ShowDialog();
                }*/
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void cobDelimiter1_SelectedIndexChanged(object sender, EventArgs e)
        {
            saveConfig();
        }

        private void cobDelimiter2_SelectedIndexChanged(object sender, EventArgs e)
        {
            saveConfig();
        }

        private void cbHeader1_CheckedChanged(object sender, EventArgs e)
        {
            saveConfig();
        }

        private void saveConfig()
        {
            if (initFinished)
            {
                Properties.Settings.Default.header1 = !!cbHeader1.Checked;
                Properties.Settings.Default.header2 = !!cbHeader2.Checked;
                Properties.Settings.Default.filename1 = !!cbUseFileName1.Checked;
                Properties.Settings.Default.filename2 = !!cbUseFileName2.Checked;
                Properties.Settings.Default.delimiter1 = (int)cobDelimiter1.SelectedIndex;
                Properties.Settings.Default.delimiter2 = (int)cobDelimiter2.SelectedIndex;
                Properties.Settings.Default.delimiter3 = (int)cobDelimiter3.SelectedIndex;
                Properties.Settings.Default.id1 = (int)numCad.Value;
                Properties.Settings.Default.id2 = (int)numBom.Value;
                Properties.Settings.Default.column1 = txtColumn1.Text;
                Properties.Settings.Default.column2 = txtColumn2.Text;
                Properties.Settings.Default.output = llblOutputDir.Text;
                Properties.Settings.Default.encoding = cobEncoding.SelectedIndex;
                Properties.Settings.Default.Save();
            }
        }

        private void cbUseFileName1_CheckedChanged(object sender, EventArgs e)
        {
            saveConfig();
            txtColumn1.Enabled = this.cbUseFileName1.Checked;

        }

        private void cbHeader2_CheckedChanged(object sender, EventArgs e)
        {
            saveConfig();
        }

        private void cbUseFileName2_CheckedChanged(object sender, EventArgs e)
        {
            saveConfig();
            txtColumn2.Enabled = this.cbUseFileName2.Checked;
        }

        private void numCad_ValueChanged(object sender, EventArgs e)
        {
            saveConfig();
        }

        private void numBom_ValueChanged(object sender, EventArgs e)
        {
            saveConfig();
        }

        private void txtColumn1_Leave(object sender, EventArgs e)
        {
            saveConfig();
        }

        private void txtColumn2_Leave(object sender, EventArgs e)
        {
            saveConfig();
        }

        private void cobEncoding_SelectedIndexChanged(object sender, EventArgs e)
        {
            saveConfig();
        }

        private void Loading(bool value)
        {
            if (value)
            {
                // disable
                toolStripStatusLabel1.Text = "Loading...";
                this.Enabled = false;
            }
            else
            {
                // enable
                toolStripStatusLabel1.Text = "";
                this.Enabled = true;
            }
        }

        private void cobDelimiter3_SelectedIndexChanged(object sender, EventArgs e)
        {
            saveConfig();
        }
    }
}

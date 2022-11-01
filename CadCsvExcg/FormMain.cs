using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace CadCsvExcg
{
    public partial class FormMain : Form
    {

        private readonly bool initFinished = false;
        private class ComboItem
        {
            public int ID { get; set; }
            public string Text { get; set; }
        }
        public FormMain()
        {
            InitializeComponent();
            // InitializeDatabase();
            ComboItem[] delimiters = new ComboItem[] {
                new ComboItem{ ID = (int)Delimiter.COMMA, Text = "Comma (,)" },
                new ComboItem{ ID = (int)Delimiter.TAB, Text = "Tab (    )" },
            };
            ComboItem[] encoders = new ComboItem[]
            {
                new ComboItem{ ID = (int)Encoder.UTF7, Text = "UTF-7" },
                new ComboItem{ ID = (int)Encoder.UTF8, Text = "UTF-8" },
                new ComboItem{ ID = (int)Encoder.UTF16LE, Text = "UTF-16LE" },
                new ComboItem{ ID = (int)Encoder.UTF16BE, Text = "UTF-16BE" },
                new ComboItem{ ID = (int)Encoder.UTF32, Text = "UTF-32" },
            };

            this.cobDelimiter1.DataSource = delimiters.Clone();
            this.cobDelimiter1.DisplayMember = "Text";
            this.cobDelimiter1.ValueMember = "ID";
            this.cobDelimiter2.DataSource = delimiters.Clone();
            this.cobDelimiter2.DisplayMember = "Text";
            this.cobDelimiter2.ValueMember = "ID";
            this.cobDelimiter3.DataSource = delimiters.Clone();
            this.cobDelimiter3.DisplayMember = "Text";
            this.cobDelimiter3.ValueMember = "ID";
            this.cobEncoding.DataSource = encoders;
            this.cobEncoding.DisplayMember = "Text";
            this.cobEncoding.ValueMember = "ID";

            LoadConfigures();

            this.txtColumn1.Enabled = this.cbUseFileName1.Checked;
            this.txtColumn2.Enabled = this.cbUseFileName2.Checked;

            CheckCombinable();

            initFinished = true;
        }

        private void LoadConfigures()
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
            this.numQuantity.Value = (int)Properties.Settings.Default.quantity;
        }


        /*static void InitializeDatabase()
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
        }*/

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
            AddFile(this.lvCad);
        }

        private void btnAdd2_Click(object sender, EventArgs e)
        {
            AddFile(this.lvBom);
        }

        private void AddFile(ListView listView)
        {
            try
            {
                using (OpenFileDialog dialog = new OpenFileDialog()) {
                    dialog.Filter = "CSV Files (*.csv)|*.csv";
                    dialog.Multiselect = true;
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        foreach (var fileName in dialog.FileNames)
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
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ResizeListViewColumnsWidth(listView);
                CheckCombinable();
            }
        }

        private void RemoveSelectedItems(ListView listView)
        {
            if (listView.SelectedItems.Count > 0)
            {
                foreach (ListViewItem item in listView.SelectedItems)
                {
                    listView.Items.Remove(item);
                }
                ResizeListViewColumnsWidth(listView);
                CheckCombinable();
            }
        }

        private void ResizeListViewColumnsWidth(ListView listView)
        {
            listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        private void btnRemove1_Click(object sender, EventArgs e)
        {
            RemoveSelectedItems(this.lvCad);
        }

        private void btnRemove2_Click(object sender, EventArgs e)
        {
            RemoveSelectedItems(this.lvBom);
        }

        private void lvCad_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.btnRemove1.Enabled = !!(this.lvCad.SelectedItems.Count > 0);
            this.btnPreview1.Enabled = !!(this.lvCad.SelectedItems.Count == 1);
        }

        private void lvBom_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.btnRemove2.Enabled = !!(this.lvBom.SelectedItems.Count > 0);
            this.btnPreview2.Enabled = !!(this.lvBom.SelectedItems.Count == 1);
        }

        private void CheckCombinable()
        {
            this.grpSettings.Enabled = this.lvCad.Items.Count > 0 && this.lvBom.Items.Count > 0;
            this.btnStart.Enabled = this.lvCad.Items.Count > 0 && this.lvBom.Items.Count > 0;
        }

        private void btnSelectDir_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    this.llblOutputDir.Text = dialog.SelectedPath;
                    SaveConfigure();
                }
            }
        }

        private void llblOutputDir_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                string fullPath = System.IO.Path.GetFullPath((string)this.llblOutputDir.Text);
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
                Preview(lvCad.SelectedItems[0].SubItems[1].Text, 1);
            }
        }

        private void btnPreview2_Click(object sender, EventArgs e)
        {
            if (lvBom.SelectedItems.Count == 1)
            {
                Preview(lvBom.SelectedItems[0].SubItems[1].Text, 2);
            }
        }

        private void lvCad_DoubleClick(object sender, EventArgs e)
        {
            if (lvCad.SelectedItems.Count == 1)
            {
                Preview(lvCad.SelectedItems[0].SubItems[1].Text, 1);
            }
        }

        private void lvBom_DoubleClick(object sender, EventArgs e)
        {
            if (lvBom.SelectedItems.Count == 1)
            {
                Preview(lvBom.SelectedItems[0].SubItems[1].Text, 2);
            }
        }
        private void Preview(string path, int configurationNumber)
        {
            Loading(true);
            using (FormProcess frm = new FormProcess())
            {
                bool isCad = configurationNumber == 1;
                List<object> args = new List<object>
                {
                    path,
                    configurationNumber,
                    isCad
                };
                frm.Preview(args);
                frm.ShowDialog();
            }
            Loading(false);
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            using (FormProcess frm = new FormProcess())
            {
                List<string> paths1 = new List<string>();
                List<string> paths2 = new List<string>();
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
                    };

                frm.Combine(args);
                frm.ShowDialog();
            }
        }

        private void cobDelimiter1_SelectedIndexChanged(object sender, EventArgs e)
        {
            SaveConfigure();
        }

        private void cobDelimiter2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SaveConfigure();
        }

        private void cbHeader1_CheckedChanged(object sender, EventArgs e)
        {
            SaveConfigure();
        }

        private void SaveConfigure()
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
                Properties.Settings.Default.quantity = (int)numQuantity.Value;
                Properties.Settings.Default.Save();
            }
        }

        private void cbUseFileName1_CheckedChanged(object sender, EventArgs e)
        {
            SaveConfigure();
            txtColumn1.Enabled = this.cbUseFileName1.Checked;

        }

        private void cbHeader2_CheckedChanged(object sender, EventArgs e)
        {
            SaveConfigure();
        }

        private void cbUseFileName2_CheckedChanged(object sender, EventArgs e)
        {
            SaveConfigure();
            txtColumn2.Enabled = this.cbUseFileName2.Checked;
        }

        private void numCad_ValueChanged(object sender, EventArgs e)
        {
            SaveConfigure();
        }

        private void numBom_ValueChanged(object sender, EventArgs e)
        {
            SaveConfigure();
        }

        private void txtColumn1_Leave(object sender, EventArgs e)
        {
            SaveConfigure();
        }

        private void txtColumn2_Leave(object sender, EventArgs e)
        {
            SaveConfigure();
        }

        private void cobEncoding_SelectedIndexChanged(object sender, EventArgs e)
        {
            SaveConfigure();
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
            SaveConfigure();
        }

        private void numQuantity_ValueChanged(object sender, EventArgs e)
        {
            SaveConfigure();
        }
    }
}

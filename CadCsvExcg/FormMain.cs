using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace CadCsvExcg
{
    public partial class frmMain : Form
    {
        private static bool initializeFinished = false;
        private bool isLoadingConfigure = false;
        private class ComboItem
        {
            public int ID { get; set; }
            public string Text { get; set; }
        }
        public frmMain()
        {
            InitializeComponent();
            ComboItem[] delimiters = new ComboItem[] {
                new ComboItem{ ID = (int)Delimiter.COMMA, Text = "Comma ( , )" },
                new ComboItem{ ID = (int)Delimiter.TAB, Text = "Tab (   )" },
            };
            ComboItem[] encoders = new ComboItem[]
            {
                new ComboItem{ ID = (int)Encoder.UTF7, Text = "UTF-7" },
                new ComboItem{ ID = (int)Encoder.UTF8, Text = "UTF-8" },
                new ComboItem{ ID = (int)Encoder.UTF16LE, Text = "UTF-16LE" },
                new ComboItem{ ID = (int)Encoder.UTF16BE, Text = "UTF-16BE" },
                new ComboItem{ ID = (int)Encoder.UTF32, Text = "UTF-32" },
            };
            ComboItem[] cadList = new ComboItem[]
            {
                new ComboItem{ ID = 0, Text = "Unidraf" },
                new ComboItem{ ID = 1, Text = "EPLAN" },
                new ComboItem{ ID = 2, Text = "E3" },
            };
            this.ddl_cad_delimiter.DataSource = delimiters.Clone();
            this.ddl_cad_delimiter.DisplayMember = "Text";
            this.ddl_cad_delimiter.ValueMember = "ID";
            this.ddl_bom_delimiter.DataSource = delimiters.Clone();
            this.ddl_bom_delimiter.DisplayMember = "Text";
            this.ddl_bom_delimiter.ValueMember = "ID";
            this.ddl_output_delimiter.DataSource = delimiters.Clone();
            this.ddl_output_delimiter.DisplayMember = "Text";
            this.ddl_output_delimiter.ValueMember = "ID";
            this.ddl_output_encoding.DataSource = encoders;
            this.ddl_output_encoding.DisplayMember = "Text";
            this.ddl_output_encoding.ValueMember = "ID";
            this.ddl_cad_type.DataSource = cadList;
            this.ddl_cad_type.DisplayMember = "Text";
            this.ddl_cad_type.ValueMember = "ID";

            this.txt_output_include.Enabled = this.cb_output_include.Checked;

            LoadConfigures();
            CheckCombinable();
            initializeFinished = true;
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
            AddFile(this.lv_cad);
        }

        private void btnAdd2_Click(object sender, EventArgs e)
        {
            AddFile(this.lv_bom);
        }

        private void AddFile(ListView listView)
        {
            try
            {
                using (OpenFileDialog dialog = new OpenFileDialog())
                {
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
            RemoveSelectedItems(this.lv_cad);
        }

        private void btnRemove2_Click(object sender, EventArgs e)
        {
            RemoveSelectedItems(this.lv_bom);
        }

        private void lvCad_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.btn_remove1.Enabled = !!(this.lv_cad.SelectedItems.Count > 0);
            this.btn_preview1.Enabled = !!(this.lv_cad.SelectedItems.Count == 1);
        }

        private void lvBom_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.btn_remove2.Enabled = !!(this.lv_bom.SelectedItems.Count > 0);
            this.btn_preview2.Enabled = !!(this.lv_bom.SelectedItems.Count == 1);
        }

        private void CheckCombinable()
        {
            this.btn_start.Enabled = this.lv_cad.Items.Count > 0 && this.lv_bom.Items.Count > 0;
        }

        private void btnSelectDir_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    this.lbtn_output_dir.Text = dialog.SelectedPath;
                    SaveConfigure(sender, e);
                }
            }
        }

        private void llblOutputDir_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                string fullPath = System.IO.Path.GetFullPath((string)this.lbtn_output_dir.Text);
                Process.Start("explorer.exe", @"" + fullPath + "");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnPreview1_Click(object sender, EventArgs e)
        {
            if (lv_cad.SelectedItems.Count == 1)
            {
                Preview(lv_cad.SelectedItems[0].SubItems[1].Text, 1);
            }
        }

        private void btnPreview2_Click(object sender, EventArgs e)
        {
            if (lv_bom.SelectedItems.Count == 1)
            {
                Preview(lv_bom.SelectedItems[0].SubItems[1].Text, 2);
            }
        }

        private void lvCad_DoubleClick(object sender, EventArgs e)
        {
            if (lv_cad.SelectedItems.Count == 1)
            {
                Preview(lv_cad.SelectedItems[0].SubItems[1].Text, 1);
            }
        }

        private void lvBom_DoubleClick(object sender, EventArgs e)
        {
            if (lv_bom.SelectedItems.Count == 1)
            {
                Preview(lv_bom.SelectedItems[0].SubItems[1].Text, 2);
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
            if ((this.txt_output_include.Text.Length > 0 && IsDigitsOnly(this.txt_output_include.Text)) || !this.cb_output_include.Checked)
            {
                using (FormProcess frm = new FormProcess())
                {
                    List<string> paths1 = new List<string>();
                    List<string> paths2 = new List<string>();
                    foreach (ListViewItem item in lv_cad.Items)
                    {
                        string file = item.SubItems[1].Text;
                        paths1.Add(file);
                    }
                    foreach (ListViewItem item in lv_bom.Items)
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
            } else
            {
                MessageBox.Show("Invalid arguments");
            }
        }

        private void LoadConfigures(bool cadOnly = false)
        {
            isLoadingConfigure = true;
            // this.ddl_cad_type.SelectedIndex = (int)Properties.Settings.Default.cad_type;
            switch (this.ddl_cad_type.SelectedIndex)
            {
                case 0:
                    this.cb_cad_header.Checked = !!Properties.Settings.Default.cad_header1;
                    this.ddl_cad_delimiter.SelectedIndex = (int)Properties.Settings.Default.cad_delimiter1;
                    this.num_cad_id_pos.Value = (int)Properties.Settings.Default.cad_id_pos1;
                    this.num_cad_quantity_pos.Value = (int)Properties.Settings.Default.cad_quantity_pos1;
                    break;
                case 1:
                    this.cb_cad_header.Checked = !!Properties.Settings.Default.cad_header2;
                    this.ddl_cad_delimiter.SelectedIndex = (int)Properties.Settings.Default.cad_delimiter2;
                    this.num_cad_id_pos.Value = (int)Properties.Settings.Default.cad_id_pos2;
                    this.num_cad_quantity_pos.Value = (int)Properties.Settings.Default.cad_quantity_pos2;
                    break;
                case 2:
                    this.cb_cad_header.Checked = !!Properties.Settings.Default.cad_header3;
                    this.ddl_cad_delimiter.SelectedIndex = (int)Properties.Settings.Default.cad_delimiter3;
                    this.num_cad_id_pos.Value = (int)Properties.Settings.Default.cad_id_pos3;
                    this.num_cad_quantity_pos.Value = (int)Properties.Settings.Default.cad_quantity_pos3;
                    break;
            }
            if (!cadOnly)
            {
                this.cb_bom_header.Checked = !!Properties.Settings.Default.bom_header;
                this.ddl_bom_delimiter.SelectedIndex = (int)Properties.Settings.Default.bom_delimiter;
                this.num_bom_id_pos.Value = (int)Properties.Settings.Default.bom_id_pos;
                // OUTPUT
                this.lbtn_output_dir.Text = Properties.Settings.Default.output_dir;
                this.ddl_output_encoding.SelectedIndex = (int)Properties.Settings.Default.output_encoding;
                this.ddl_output_delimiter.SelectedIndex = (int)Properties.Settings.Default.output_delimiter;
                this.cb_output_exclude.Checked = !!Properties.Settings.Default.output_exclude;
                this.txt_output_include.Text = Properties.Settings.Default.output_include;
                this.cb_output_include.Checked = !!Properties.Settings.Default.output_isinclude;
                switch (Properties.Settings.Default.output_repeat)
                {
                    case 0:
                        rdo_output_option1.Checked = true;
                        rdo_output_option2.Checked = false;
                        rdo_output_option3.Checked = false;
                        break;
                    case 1:
                        rdo_output_option1.Checked = false;
                        rdo_output_option2.Checked = true;
                        rdo_output_option3.Checked = false;
                        break;
                    case 2:
                        rdo_output_option1.Checked = false;
                        rdo_output_option2.Checked = false;
                        rdo_output_option3.Checked = true;
                        break;
                }
            }
            
            isLoadingConfigure = false;
        }

        private void SaveConfigure(object sender, EventArgs e)
        {
            if (initializeFinished && !isLoadingConfigure)
            {
                switch ((int)this.ddl_cad_type.SelectedIndex)
                {
                    case 0:
                        Properties.Settings.Default.cad_header1 = !!this.cb_cad_header.Checked;
                        Properties.Settings.Default.cad_id_pos1 = (int)this.num_cad_id_pos.Value;
                        Properties.Settings.Default.cad_quantity_pos1 = (int)this.num_cad_quantity_pos.Value;
                        Properties.Settings.Default.cad_delimiter1 = (int)this.ddl_cad_delimiter.SelectedIndex;
                        break;
                    case 1:
                        Properties.Settings.Default.cad_header2 = !!this.cb_cad_header.Checked;
                        Properties.Settings.Default.cad_id_pos2 = (int)this.num_cad_id_pos.Value;
                        Properties.Settings.Default.cad_quantity_pos2 = (int)this.num_cad_quantity_pos.Value;
                        Properties.Settings.Default.cad_delimiter2 = (int)this.ddl_cad_delimiter.SelectedIndex;
                        break;
                    case 2:
                        Properties.Settings.Default.cad_header3 = !!this.cb_cad_header.Checked;
                        Properties.Settings.Default.cad_id_pos3 = (int)this.num_cad_id_pos.Value;
                        Properties.Settings.Default.cad_quantity_pos3 = (int)this.num_cad_quantity_pos.Value;
                        Properties.Settings.Default.cad_delimiter3 = (int)this.ddl_cad_delimiter.SelectedIndex;
                        break;
                }

                Properties.Settings.Default.bom_header = !!this.cb_bom_header.Checked;
                Properties.Settings.Default.bom_delimiter = (int)this.ddl_bom_delimiter.SelectedIndex;
                Properties.Settings.Default.bom_id_pos = (int)this.num_bom_id_pos.Value;

                Properties.Settings.Default.output_dir = this.lbtn_output_dir.Text;
                Properties.Settings.Default.output_encoding = (int)this.ddl_output_encoding.SelectedIndex;
                Properties.Settings.Default.output_delimiter = (int)this.ddl_output_delimiter.SelectedIndex;
                Properties.Settings.Default.output_exclude = !!this.cb_output_exclude.Checked;
                Properties.Settings.Default.output_include = this.txt_output_include.Text;
                Properties.Settings.Default.output_isinclude = this.cb_output_include.Checked;
                if (rdo_output_option1.Checked)
                {
                    Properties.Settings.Default.output_repeat = 0;
                } else if (rdo_output_option2.Checked)
                {
                    Properties.Settings.Default.output_repeat = 1;
                } else if (rdo_output_option3.Checked)
                {
                    Properties.Settings.Default.output_repeat = 2;
                }
                // Properties.Settings.Default.cad_type = (int)this.ddl_cad_type.SelectedIndex;
                Properties.Settings.Default.Save();
            }
        }

        private void Loading(bool value)
        {
            if (value)
            {
                // disable
                toolStripStatusLabel1.Text = "読み込み中";
                this.Enabled = false;
            }
            else
            {
                // enable
                toolStripStatusLabel1.Text = "";
                this.Enabled = true;
            }
        }

        private void ddl_cad_type_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadConfigures(true);
        }

        private void cb_output_include_CheckedChanged(object sender, EventArgs e)
        {
            this.txt_output_include.Enabled = this.cb_output_include.Checked;
            SaveConfigure(sender, e);
        }

        bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                    return false;
            }

            return true;
        }
    }
}

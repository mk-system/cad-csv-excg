using System.Data.SQLite;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace CadCsvExcg
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            InitializeDatabase();
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
                }
                finally
                {
                    connection.Close();
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

        private void cbUseFileName_CheckedChanged(object sender, EventArgs e)
        {
            numCad.Enabled = !cbUseFileName.Checked;
            lblCad.Enabled = numCad.Enabled;
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
                            string[] arr = new string[4];
                            arr[0] = System.IO.Path.GetFileName(fileName);
                            arr[1] = fileName;
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
            Preview(lvCad.SelectedItems[0].SubItems[1].Text);
        }

        private void Preview(string path)
        {
            if (File.Exists(path))
            {
                var frm = new frmPreview();
                frm.path = path;
                frm.ShowDialog(this);
            }
            
        }

        private void lvCad_DoubleClick(object sender, EventArgs e)
        {
            Preview(lvCad.SelectedItems[0].SubItems[1].Text);
        }
    }
}

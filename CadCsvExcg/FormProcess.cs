using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CadCsvExcg
{
    public partial class FormProcess : Form
    {
        public FormProcess()
        {
            InitializeComponent();
            this.ControlBox = false;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;

            backgroundWorker2.WorkerReportsProgress = true;
            backgroundWorker2.WorkerSupportsCancellation = true;

            button1.Text = "Cancel";
        }

        public void Preview(List<object> args)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                button1.Click += new EventHandler(this.CancelBackgroundWorker1);
                backgroundWorker1.RunWorkerAsync(args);
            }
        }

        public void Combine(List<object> args)
        {
            if (backgroundWorker2.IsBusy != true)
            {
                button1.Click += new EventHandler(this.CancelBackgroundWorker2);
                backgroundWorker2.RunWorkerAsync(args);
            }
        }


        private void CancelBackgroundWorker1(object sender, EventArgs e)
        {
            if (backgroundWorker1.WorkerSupportsCancellation == true)
            {
                backgroundWorker1.CancelAsync();
            }
        }

        private void CancelBackgroundWorker2(object sender, EventArgs e)
        {
            if (backgroundWorker2.WorkerSupportsCancellation == true)
            {
                backgroundWorker2.CancelAsync();
            }
        }


        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                DataTable dt = new DataTable();
                List<object> genericList = e.Argument as List<object>;
                string path = (string)genericList[0];
                string delimiter = (string)genericList[1];
                bool is_first_line_header = (bool)genericList[2];
                int column_start_number = (int)genericList[3];
                int primary_column = (int)genericList[4];
                bool extra_column = (bool)genericList[5];
                string extra_column_name = (string)genericList[6];
                bool is_cad = (bool)genericList[7];
                int quantity_column = (int)genericList[8];

                string file_name = Path.GetFileNameWithoutExtension(path);
                int numberOfLines = File.ReadLines(path).Count();
                int lineCount = 0;
                using (TextFieldParser reader = new TextFieldParser(path))
                {
                    reader.TextFieldType = FieldType.Delimited;
                    reader.SetDelimiters(new string[] { delimiter });
                    reader.HasFieldsEnclosedInQuotes = true;
                    string[] colFields = reader.ReadFields();

                    if (is_cad && primary_column > 0 && quantity_column > 0)
                    {
                        colFields = colFields.Where((el, i) => i == primary_column - 1 || i == quantity_column - 1).ToArray();
                    }

                    if (extra_column)
                    {
                        if (is_first_line_header)
                        {
                            colFields = PrependNewField(colFields, extra_column_name);
                        }
                        else
                        {
                            colFields = PrependNewField(colFields, file_name);
                        }
                    }

                    foreach (var (column, i) in colFields.Select((v, i) => (v, i)))
                    {
                        DataColumn dc = new DataColumn(is_first_line_header ? column : ("COLUMN " + (column_start_number + i + 1)));
                        dc.AllowDBNull = true;
                        dt.Columns.Add(dc);
                    }

                    if (!is_first_line_header)
                    {
                        dt.Rows.Add(colFields);
                    }

                    lineCount++;

                    while (!reader.EndOfData)
                    {
                        if (worker.CancellationPending == true)
                        {
                            e.Cancel = true;
                            break;
                        }
                        else
                        {
                            string[] fieldData = reader.ReadFields();

                            if (is_cad && primary_column > 0 && quantity_column > 0)
                            {
                                fieldData = fieldData.Where((el, i) => i == primary_column - 1 || i == quantity_column - 1).ToArray();
                            }

                            if (extra_column)
                            {
                                fieldData = PrependNewField(fieldData, file_name);
                            }

                            for (int i = 0; i < fieldData.Length; i++)
                            {
                                if (fieldData[i] == "")
                                {
                                    fieldData[i] = null;
                                }
                            }
                            dt.Rows.Add(fieldData);
                            lineCount++;

                            int complete = (int)Math.Round((double)(100 * lineCount) / numberOfLines);
                            if (this.progressBar1.Value < complete)
                            {
                                worker.ReportProgress(complete);
                            }

                        }
                    }
                }
                e.Result = dt;
            }
            catch (Exception)
            { }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            lblResult.Text = (e.ProgressPercentage.ToString() + "%");
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                lblResult.Text = "Canceled!";

            }
            else if (e.Error != null)
            {
                lblResult.Text = "Error: " + e.Error.Message;
            }
            else
            {
                progressBar1.Value = progressBar1.Maximum;
                lblResult.Text = "Done!";
                button1.Enabled = false;
                using (FormPreview frm = new FormPreview((DataTable)e.Result))
                {
                    frm.ShowDialog();
                }
            }
            this.Close();
        }


        private static string[] PrependNewField(string[] target, string data)
        {
            List<string> list = new List<string>(target);
            list.Insert(0, data);
            return list.ToArray();
        }

        private static string[] BuildNewFields(string[] target, int primary, int quantity)
        {
            List<string> list = new List<string>(target);
            return list.ToArray();
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                BackgroundWorker worker = sender as BackgroundWorker;
                DataTable dt1 = new DataTable();
                DataTable dt2 = new DataTable();
                DataTable dt3 = new DataTable();
                List<object> genericList = e.Argument as List<object>;
                string[] paths1 = (string[])genericList[0];
                string[] paths2 = (string[])genericList[1];
                int total = paths1.Count() + paths2.Count();
                int current = 0;

                // Repeat twice; 0: cad, 1: bom
                for (var i = 0; i < 2; i++)
                {
                    int column_start_number = 0;
                    if (i != 0)
                    {
                        column_start_number = (int)dt1.Columns.Count;
                    }
                    foreach (var path in (string[])genericList[0 + i])
                    {
                        DataTable dt = new DataTable();

                        string file_name = Path.GetFileNameWithoutExtension(path);

                        int lineCount = 0;

                        int numberOfLines = File.ReadLines(path).Count();

                        using (TextFieldParser reader = new TextFieldParser(path))
                        {
                            reader.TextFieldType = FieldType.Delimited;
                            reader.SetDelimiters(new string[] { (string)genericList[2 + i] });
                            reader.HasFieldsEnclosedInQuotes = true;
                            string[] colFields = reader.ReadFields();

                            if (i == 0 && (int)genericList[6] > 0 && (int)genericList[12] > 0)
                            {
                                colFields = colFields.Where((el, id) => id == (int)genericList[6] - 1 || id == (int)genericList[12] - 1).ToArray();
                            }

                            if ((bool)genericList[8 + i])
                            {
                                if ((bool)genericList[4 + i])
                                {
                                    colFields = PrependNewField(colFields, (string)genericList[10 + i]);
                                }
                                else
                                {
                                    colFields = PrependNewField(colFields, file_name);
                                }
                            }

                            foreach (var (column, j) in colFields.Select((v, j) => (v, j)))
                            {
                                DataColumn dc = new DataColumn((bool)genericList[4 + i] ? column : ("COLUMN " + (column_start_number + j + 1)));
                                dc.AllowDBNull = true;
                                dt.Columns.Add(dc);
                            }

                            if (!(bool)genericList[4 + i])
                            {
                                dt.Rows.Add(colFields);
                            }
                            while (!reader.EndOfData)
                            {
                                if (worker.CancellationPending == true)
                                {
                                    e.Cancel = true;
                                    break;
                                }
                                else
                                {
                                    string[] fieldData = reader.ReadFields();

                                    if (i == 0 && (int)genericList[6] > 0 && (int)genericList[12] > 0)
                                    {
                                        fieldData = fieldData.Where((el, id) => id == (int)genericList[6] - 1 || id == (int)genericList[12] - 1).ToArray();
                                    }

                                    if ((bool)genericList[8 + i])
                                    {
                                        fieldData = PrependNewField(fieldData, file_name);
                                    }

                                    for (int j = 0; j < fieldData.Length; j++)
                                    {
                                        if (fieldData[j] == "")
                                        {
                                            fieldData[j] = null;
                                        }
                                    }
                                    dt.Rows.Add(fieldData);

                                    lineCount++;

                                    int complete = (int)Math.Round((double)(100 * lineCount) / numberOfLines / total + (100 * current - 1) / total);

                                    if (this.progressBar1.Value < complete)
                                    {
                                        worker.ReportProgress(complete);
                                    }
                                }
                            }
                        }

                        if (i == 0)
                        {
                            dt1.Merge(dt);
                        }
                        else
                        {
                            dt2.Merge(dt);
                        }
                        current++;
                    }
                }

                int primary1 = (bool)genericList[8] ? (int)genericList[6] : (int)genericList[6] - 1;
                int primary2 = (bool)genericList[9] ? (int)genericList[7] : (int)genericList[7] - 1;

                // getting all columns
                dt3.Merge(dt1, false, MissingSchemaAction.AddWithKey);
                dt3.Merge(dt2, false, MissingSchemaAction.Add);
                DataTable resultDt = new DataTable();
                resultDt = dt3.Clone();
                // joining datatables
                var result = from table1 in dt1.AsEnumerable()
                            join table2 in dt2.AsEnumerable() on new { ID = table1[primary1] } equals new { ID = table2[primary2] }
                            into temp
                            from table3 in temp.DefaultIfEmpty()
                            select resultDt.LoadDataRow(Concatenate((object[])table1.ItemArray, (object[])table3.ItemArray, primary2), false);

                  e.Result = result.CopyToDataTable();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            lblResult.Text = (e.ProgressPercentage.ToString() + "%");
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                lblResult.Text = "Canceled!";

            }
            else if (e.Error != null)
            {
                lblResult.Text = "Error: " + e.Error.Message;
            }
            else
            {
                progressBar1.Value = progressBar1.Maximum;
                lblResult.Text = "Done!";
                button1.Enabled = false;
                DateTime now = DateTime.Now;
                string outputPath = Properties.Settings.Default.output + "\\" + now.ToString("ddMMyyyyHHmmss") + ".csv";
                bool header = Properties.Settings.Default.header1 || Properties.Settings.Default.header2;
                DataTable dt = (DataTable)e.Result;
                string delimiter = Properties.Settings.Default.delimiter3 == 0 ? Delimiter.COMMA.GetString() : Delimiter.TAB.GetString();

                Encoding encoding;
                switch ((int)Properties.Settings.Default.encoding)
                {
                    case 0:
                        encoding = Encoding.UTF7;
                        break;
                    case 1:
                        encoding = Encoding.UTF8;
                        break;
                    case 2:
                        encoding = Encoding.Unicode;
                        break;
                    case 3:
                        encoding = Encoding.BigEndianUnicode;
                        break;
                    case 4:
                        encoding = Encoding.UTF32;
                        break;
                    default:
                        encoding = Encoding.UTF8;
                        break;
                }

                try { 
                    using (Stream s = File.Create(outputPath))
                    {
                        StreamWriter sw = new StreamWriter(s, encoding);
                        if (header)
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                sw.Write(dt.Columns[i]);
                                if (i < dt.Columns.Count - 1)
                                {
                                    sw.Write(delimiter);
                                }
                            }
                            sw.Write(sw.NewLine);
                        }
                        foreach (DataRow dr in dt.Rows)
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                if (!Convert.IsDBNull(dr[i]))
                                {
                                    string value = dr[i].ToString();
                                    if (value.Contains(","))
                                    {
                                        value = String.Format("\"{0}\"", value);
                                        sw.Write(value);
                                    }
                                    else
                                    {
                                        sw.Write(dr[i].ToString());
                                    }
                                }
                                if (i < dt.Columns.Count - 1)
                                {
                                    sw.Write(delimiter);
                                }
                            }
                            sw.Write(sw.NewLine);
                        }
                        sw.Close();
                    }
                } catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                try
                {
                    var fullPath = System.IO.Path.GetFullPath(Properties.Settings.Default.output);
                    Process.Start("explorer.exe", @"" + fullPath + "");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            this.Close();
        }

        public static T[] Concatenate<T>(T[] array1, T[] array2, int excludedIndex = 0)
        {
            T[] result = new T[array1.Length + array2.Length - 1];
            array1.CopyTo(result, 0);
            array2.Where((v, i) => i != excludedIndex).ToArray().CopyTo(result, array1.Length);
            return result;
        }
    }

}

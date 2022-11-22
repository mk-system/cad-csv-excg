﻿using Microsoft.VisualBasic.FileIO;
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

                List<object> genericList = e.Argument as List<object>;
                string path = (string)genericList[0];
                int configurationNumber = (int)genericList[1];
                bool isCAD = (bool)genericList[2];
                CSVConfig config = new CSVConfig(configurationNumber);

                DataTable dt = new DataTable();

                string fileName = Path.GetFileNameWithoutExtension(path);
                long totalLines = File.ReadLines(path).Count();

                using (TextFieldParser reader = new TextFieldParser(path))
                {
                    reader.TextFieldType = FieldType.Delimited;
                    reader.SetDelimiters(new string[] { config.delimiter });
                    reader.HasFieldsEnclosedInQuotes = true;

                    long currentLine = 0;

                    while (!reader.EndOfData)
                    {
                        if (worker.CancellationPending == true)
                        {
                            e.Cancel = true;
                            break;
                        }
                        else
                        {
                            string[] fields = reader.ReadFields();

                            if (isCAD && config.columnMatchPosition > 0 && config.columnQuantityPosition > 0)
                            {
                                fields = fields.Where((el, i) => i == config.columnMatchPosition - 1 || i == config.columnQuantityPosition - 1).ToArray();
                            }
                            if (config.usingFilename)
                            {
                                string newField;
                                if (currentLine == 0 && config.hasHeaderRow)
                                {
                                    newField = config.columnFilenameText;
                                }
                                else
                                {
                                    newField = fileName;
                                }
                                fields = PrependNewField(fields, newField);
                            }
                            for (int i = 0; i < fields.Length; i++)
                            {
                                if (fields[i] == "")
                                {
                                    fields[i] = null;
                                }
                            }
                            if (currentLine == 0)
                            {
                                bool duplicated = false;
                                if (config.hasHeaderRow && fields.Length != fields.Distinct().Count())
                                {
                                    MessageBox.Show("Could not set first line as header.", "Duplicated fields");
                                    duplicated = true;
                                }
                                foreach (var field in fields)
                                {
                                    DataColumn dc = new DataColumn(config.hasHeaderRow && !duplicated ? field : null)
                                    {
                                        AllowDBNull = true
                                    };
                                    dt.Columns.Add(dc);
                                }
                                if (!config.hasHeaderRow || duplicated)
                                {
                                    dt.Rows.Add(fields);
                                }
                            }
                            else
                            {
                                dt.Rows.Add(fields);
                            }
                            currentLine = reader.LineNumber;
                            int complete = (int)Math.Round((double)(100 * currentLine) / totalLines);
                            // only update progressbar if complete is increased
                            if (this.progressBar1.Value < complete)
                            {
                                worker.ReportProgress(complete);
                            }
                        }

                    }
                }
                e.Result = dt;
            }
            catch (Exception ex)
            {
                if (backgroundWorker1.WorkerSupportsCancellation == true)
                {
                    backgroundWorker1.CancelAsync();
                }
                MessageBox.Show(ex.Message);
                throw;
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            lblResult.Text = (e.ProgressPercentage.ToString() + "%");
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = progressBar1.Maximum;
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

                CSVConfig config1 = new CSVConfig(1);
                CSVConfig config2 = new CSVConfig(2);

                int outputType = Properties.Settings.Default.output_repeat;

                int primary1_pos, primary2_pos, quantity_pos = 0;
                if (config1.columnQuantityPosition > 0)
                {
                    if (config1.columnQuantityPosition > config1.columnMatchPosition)
                    {
                        primary1_pos = config1.usingFilename ? 1 : 0;
                        quantity_pos = config1.usingFilename ? 2 : 1;
                    }
                    else
                    {
                        primary1_pos = config1.usingFilename ? 2 : 1;
                        quantity_pos = config1.usingFilename ? 1 : 0;
                    }
                }
                else
                {
                    primary1_pos = config1.usingFilename ? config1.columnMatchPosition : config1.columnMatchPosition - 1;
                }
                primary2_pos = config2.usingFilename ? config2.columnMatchPosition : config2.columnMatchPosition - 1;

                int totalPaths = paths1.Count() + paths2.Count();
                int currentPath = 0;
                
                int repeat = 1;
                if (paths2.Count() > 0)
                {
                    repeat = 2;
                }
                // Repeat twice; 0 -> cad, 1 -> bom
                for (int i = 0; i < repeat; i++)
                {
                    CSVConfig config = (i == 0) ? config1 : config2;
                    int columnStartNumber = 0;
                    if (i == 1)
                    {
                        columnStartNumber = (int)dt1.Columns.Count;
                    }
                    foreach (var path in (string[])genericList[0 + i])
                    {
                        DataTable dt = new DataTable();
                        string fileName = Path.GetFileNameWithoutExtension(path);
                        using (TextFieldParser reader = new TextFieldParser(path))
                        {
                            reader.TextFieldType = FieldType.Delimited;
                            reader.SetDelimiters(new string[] { config.delimiter });
                            reader.HasFieldsEnclosedInQuotes = true;
                            long totalLines = File.ReadLines(path).Count();
                            long currentLine = 0;

                            while (!reader.EndOfData)
                            {
                                if (worker.CancellationPending == true)
                                {
                                    e.Cancel = true;
                                    break;
                                }
                                else
                                {
                                    string[] fields = reader.ReadFields();
                                    if (i == 0 && config.columnMatchPosition > 0 && config.columnQuantityPosition > 0)
                                    {
                                        fields = fields.Where((el, j) => j == config.columnMatchPosition - 1 || j == config.columnQuantityPosition - 1).ToArray();
                                    }
                                    if (config.usingFilename)
                                    {
                                        string newField;
                                        if (currentLine == 0 && config.hasHeaderRow)
                                        {
                                            newField = config.columnFilenameText;
                                        }
                                        else
                                        {
                                            newField = fileName;
                                        }
                                        fields = PrependNewField(fields, newField);
                                    }
                                    for (int j = 0; j < fields.Length; j++)
                                    {
                                        if (fields[j] == "")
                                        {
                                            fields[j] = null;
                                        }
                                    }
                                    if (currentLine == 0)
                                    {
                                        bool duplicated = false;
                                        if (config.hasHeaderRow && fields.Length != fields.Distinct().Count())
                                        {
                                            MessageBox.Show("Could not set first line as header.", "Duplicated fields");
                                            duplicated = true;
                                        }
                                        foreach (var (field, j) in fields.Select((v, j) => (v, j + 1)))
                                        {
                                            string columnText = "Column";
                                            if (j == 1 && i == 0)
                                            {
                                                columnText = "親品目番号";
                                            }
                                            else if (j == primary1_pos + 1 && i == 0)
                                            {
                                                columnText = "子品目番号";
                                            } else if (j == quantity_pos + 1 && i == 0)
                                            {
                                                columnText = "員数";
                                            } else
                                            {
                                                columnText += (columnStartNumber + j);
                                            }
                                            DataColumn dc = new DataColumn(config.hasHeaderRow && !duplicated ? field : columnText)
                                            {
                                                AllowDBNull = true
                                            };
                                            dt.Columns.Add(dc);
                                        }
                                        if (i == 0 && config.usingFilename)
                                        {
                                            string[] tempField = { "", fileName, "1" };
                                            dt.Rows.Add(tempField);

                                        }
                                        if (!config.hasHeaderRow || duplicated)
                                        {
                                            dt.Rows.Add(fields);
                                        }
                                    }
                                    else
                                    {
                                        if (i == 0 && outputType == 2 && Convert.ToInt32(fields[quantity_pos]) > 1)
                                        {
                                            string[] extractedFields = new string[fields.Length];
                                            fields.CopyTo(extractedFields, 0);
                                            extractedFields[quantity_pos] = "1";
                                            for (int x = 0; x < Convert.ToInt32(fields[quantity_pos]); x++)
                                            {
                                                dt.Rows.Add(extractedFields);
                                            }
                                        } else
                                        {
                                            dt.Rows.Add(fields);
                                        }
                                    }
                                    currentLine = reader.LineNumber;
                                    int complete = (int)Math.Round((double)(100 * currentLine) / totalLines / totalPaths + (100 * currentPath - 1) / totalPaths);
                                    // only update progressbar if complete is increased
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
                        currentPath++;
                    }
                }

                if (outputType == 1)
                {
                    DataTable tmp = dt1.AsEnumerable().GroupBy(r => r[primary1_pos]).Select(g => {
                        var row = g.First();
                        row.SetField(dt1.Columns[quantity_pos].ColumnName.ToString(), g.Sum(r => Convert.ToInt32(r[quantity_pos])).ToString());
                        return row;
                    }).CopyToDataTable();
                    dt1.Clear();
                    dt1.Merge(tmp, false, MissingSchemaAction.Add);
                }

                dt3.Merge(dt1, false, MissingSchemaAction.Add);

                if (paths2.Count() > 0)
                {
                    dt3.Merge(dt2, false, MissingSchemaAction.Add);
                }

                DataTable resultDt = new DataTable();
                resultDt = dt3.Clone();
                
                if (!!Properties.Settings.Default.output_exclude)
                {
                    e.Result = dt1;
                } else
                {
                    var result = from t1 in dt1.AsEnumerable()
                                 join t2 in dt2.AsEnumerable()
                                 on new { ID = t1[primary1_pos] } equals new { ID = t2[primary2_pos] }
                                 select resultDt.LoadDataRow(Concatenate((object[])t1.ItemArray, (object[])t2.ItemArray, primary2_pos), true);
                    e.Result = result.CopyToDataTable();
                }

            }
            catch (Exception ex)
            {
                if (backgroundWorker2.WorkerSupportsCancellation == true)
                {
                    backgroundWorker2.CancelAsync();
                }
                MessageBox.Show(ex.Message);
                // throw;
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
                MessageBox.Show("CANCELED");

            }
            else if (e.Error != null)
            {
                lblResult.Text = "Error: " + e.Error.Message;
                MessageBox.Show("ERROR");
            }
            else
            {
                progressBar1.Value = progressBar1.Maximum;
                lblResult.Text = "Done!";
                button1.Enabled = false;
                DateTime now = DateTime.Now;
                string outputPath = Properties.Settings.Default.output_dir + "\\" + now.ToString("ddMMyyyyHHmmss") + ".csv";
                bool header = Properties.Settings.Default.cad_header1 || Properties.Settings.Default.bom_header;
                DataTable dt = (DataTable)e.Result;
                string delimiter = Properties.Settings.Default.output_delimiter == 0 ? Delimiter.COMMA.GetString() : Delimiter.TAB.GetString();

                Encoding encoding;
                switch ((int)Properties.Settings.Default.output_encoding)
                {
                    case (int)Encoder.UTF7:
                        encoding = Encoding.UTF7;
                        break;
                    case (int)Encoder.UTF8:
                        encoding = Encoding.UTF8;
                        break;
                    case (int)Encoder.UTF16LE:
                        encoding = Encoding.Unicode;
                        break;
                    case (int)Encoder.UTF16BE:
                        encoding = Encoding.BigEndianUnicode;
                        break;
                    case (int)Encoder.UTF32:
                        encoding = Encoding.UTF32;
                        break;
                    default:
                        encoding = Encoding.UTF8;
                        break;
                }

                try
                {
                    using (Stream s = File.Create(outputPath))
                    {
                        StreamWriter sw = new StreamWriter(s, encoding);
                        if (true)
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
                    MessageBox.Show("File saved to the output directory.", "Finished");
                    var fullPath = System.IO.Path.GetFullPath(Properties.Settings.Default.output_dir);
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

    public class CSVConfig
    {
        public string delimiter;
        public bool hasHeaderRow;
        public bool usingFilename;
        public string columnFilenameText;
        public int columnMatchPosition;
        public int columnQuantityPosition;

        public CSVConfig(int configurationNumber)
        {
            switch (configurationNumber)
            {
                case 1:
                    int type = Properties.Settings.Default.cad_type;
                    switch (type)
                    {
                        case 0:
                            this.delimiter = ((Delimiter)Properties.Settings.Default.cad_delimiter1).GetString();
                            this.hasHeaderRow = !!Properties.Settings.Default.cad_header1;
                            this.usingFilename = !!Properties.Settings.Default.cad_filename1;
                            this.columnFilenameText = Properties.Settings.Default.cad_file_column_text1;
                            this.columnMatchPosition = (int)Properties.Settings.Default.cad_id_pos1;
                            this.columnQuantityPosition = (int)Properties.Settings.Default.cad_quantity_pos1;
                            break;
                        case 1:
                            this.delimiter = ((Delimiter)Properties.Settings.Default.cad_delimiter2).GetString();
                            this.hasHeaderRow = !!Properties.Settings.Default.cad_header2;
                            this.usingFilename = !!Properties.Settings.Default.cad_filename2;
                            this.columnFilenameText = Properties.Settings.Default.cad_file_column_text2;
                            this.columnMatchPosition = (int)Properties.Settings.Default.cad_id_pos2;
                            this.columnQuantityPosition = (int)Properties.Settings.Default.cad_quantity_pos2;
                            break;
                        case 2:
                            this.delimiter = ((Delimiter)Properties.Settings.Default.cad_delimiter3).GetString();
                            this.hasHeaderRow = !!Properties.Settings.Default.cad_header3;
                            this.usingFilename = !!Properties.Settings.Default.cad_filename3;
                            this.columnFilenameText = Properties.Settings.Default.cad_file_column_text3;
                            this.columnMatchPosition = (int)Properties.Settings.Default.cad_id_pos3;
                            this.columnQuantityPosition = (int)Properties.Settings.Default.cad_quantity_pos3;
                            break;
                    }
                    break;
                case 2:
                    this.delimiter = ((Delimiter)Properties.Settings.Default.bom_delimiter).GetString();
                    this.hasHeaderRow = !!Properties.Settings.Default.bom_header;
                    this.usingFilename = false;
                    this.columnFilenameText = "";
                    this.columnMatchPosition = (int)Properties.Settings.Default.bom_id_pos;
                    this.columnQuantityPosition = -1;
                    break;
            }
        }
    }
}

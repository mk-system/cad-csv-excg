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
                                    newField = "親品目番号";
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

            BackgroundWorker worker = sender as BackgroundWorker;
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();

            List<object> genericList = e.Argument as List<object>;
            string[] paths1 = (string[])genericList[0];
            string[] paths2 = (string[])genericList[1];

            CSVConfig config1 = new CSVConfig(1);
            CSVConfig config2 = new CSVConfig(2);

            bool exclude = !!Properties.Settings.Default.output_exclude;
            int outputType = Properties.Settings.Default.output_repeat;
            string include = Properties.Settings.Default.output_include;

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
            primary2_pos = config2.columnMatchPosition - 1;

            int totalPaths = paths1.Count() + paths2.Count();
            int currentPath = 0;

            int lastColumnIndex = 0;
            // CAD files
            foreach (var path in (string[])genericList[0])
            {
                DataTable dt = new DataTable();
                string filename = Path.GetFileNameWithoutExtension(path);
                using (TextFieldParser reader = new TextFieldParser(path))
                {
                    reader.TextFieldType = FieldType.Delimited;
                    reader.SetDelimiters(new string[] { config1.delimiter });
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
                            // setting empty fields to null
                            for (int i = 0; i < fields.Length; i++)
                            {
                                if (fields[i] == "")
                                {
                                    fields[i] = null;
                                }
                            }
                            // getting 2 columns (id and quantity)
                            fields = fields.Where((_, i) => i == config1.columnMatchPosition - 1 || i == config1.columnQuantityPosition - 1).ToArray();

                            // prepending filename as column
                            if (config1.usingFilename)
                            {
                                if (currentLine == 0 && config1.hasHeaderRow)
                                {
                                    fields = PrependNewField(fields, (string)config1.columnFilenameText);
                                }
                                else
                                {
                                    fields = PrependNewField(fields, (string)filename);
                                }
                            }
                            // setting header
                            if (currentLine == 0)
                            {
                                foreach (var (field, i) in fields.Select((v, i) => (v, i + 1)))
                                {
                                    string columnText = "桁";
                                    int primaryPos = primary1_pos + 1;
                                    int quantityPos = quantity_pos + 1;
                                    if (i == 1)
                                    {
                                        columnText = "親品目番号";
                                    }
                                    else if (i == primaryPos)
                                    {
                                        columnText = "子品目番号";
                                    }
                                    else if (i == quantityPos)
                                    {
                                        columnText = "員数";
                                    }
                                    else
                                    {
                                        columnText += i;
                                        lastColumnIndex = i;
                                    }
                                    DataColumn dc = new DataColumn(columnText)
                                    {
                                        AllowDBNull = true
                                    };
                                    dt.Columns.Add(dc);
                                }

                                if (!config1.hasHeaderRow)
                                {
                                    dt.Rows.Add(fields);
                                }
                            }
                            else
                            {
                                switch (outputType)
                                {
                                    case 2:
                                        var quantity = Convert.ToInt32(fields[quantity_pos]);
                                        if (quantity > 1)
                                        {
                                            string[] newFields = new string[fields.Length];
                                            fields.CopyTo(newFields, 0);
                                            newFields[quantity_pos] = "1";
                                            for (int i = 0; i < quantity; i++)
                                            {
                                                dt.Rows.Add(newFields);
                                            }
                                        }
                                        else
                                        {
                                            dt.Rows.Add(fields);
                                        }
                                        break;
                                    default:
                                        dt.Rows.Add(fields);
                                        break;
                                }
                            }
                            currentLine = reader.LineNumber;
                            int complete = (int)Math.Round((double)(100 * currentLine) / totalLines / totalPaths + (100 * currentPath - 1) / totalPaths);
                            if (this.progressBar1.Value < complete)
                            {
                                worker.ReportProgress(complete);
                            }
                        }
                    }
                }
                dt1.Merge(dt);
                currentPath++;
            }
            // merge duplicated rows of CAD
            if (outputType == 1)
            {
                DataTable tmpDt = dt1.AsEnumerable().GroupBy(r => new
                {
                    fname = r[0],
                    match = r[primary1_pos]
                }).Select(g =>
                {
                    var row = g.First();
                    row.SetField(dt1.Columns[quantity_pos].ColumnName.ToString(), g.Sum(r => Convert.ToInt32(r[quantity_pos])).ToString());
                    return row;
                }).CopyToDataTable();
                dt1.Clear();
                dt1.Merge(tmpDt, false, MissingSchemaAction.Add);
            }
            // BOM files
            foreach (var path in (string[])genericList[1])
            {
                DataTable dt = new DataTable();
                string filename = Path.GetFileNameWithoutExtension(path);
                using (TextFieldParser reader = new TextFieldParser(path))
                {
                    reader.TextFieldType = FieldType.Delimited;
                    reader.SetDelimiters(new string[] { config2.delimiter });
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
                            for (int i = 0; i < fields.Length; i++)
                            {
                                if (fields[i] == "")
                                {
                                    fields[i] = null;
                                }
                            }
                            if (exclude == true)
                            {
                                fields = fields.Where((_, i) => i == 0 || i == primary2_pos).ToArray();
                            }

                            if (currentLine == 0)
                            {
                                bool duplicated = false;
                                if (config2.hasHeaderRow && fields.Length != fields.Distinct().Count())
                                {
                                    duplicated = true;
                                }

                                foreach (var (field, i) in fields.Select((v, i) => (v, i + 1)))
                                {
                                    string columnText = "桁";
                                    if (i == (exclude ? 2 : primary2_pos + 1))
                                    {
                                        columnText = "子品目番号";
                                    }
                                    else
                                    {
                                        columnText += (lastColumnIndex + i);
                                    }
                                    DataColumn dc = new DataColumn(config2.hasHeaderRow && !duplicated && i != (exclude ? 2 : primary2_pos + 1) ? field : columnText)
                                    {
                                        AllowDBNull = true
                                    };
                                    dt.Columns.Add(dc);
                                }
                                if (!config2.hasHeaderRow)
                                {
                                    dt.Rows.Add(fields);
                                }
                            }
                            else
                            {
                                dt.Rows.Add(fields);
                            }

                            currentLine = reader.LineNumber;
                            int complete = (int)Math.Round((double)(100 * currentLine) / totalLines / totalPaths + (100 * currentPath - 1) / totalPaths);
                            if (this.progressBar1.Value < complete)
                            {
                                worker.ReportProgress(complete);
                            }
                        }
                    }
                }
                dt2.Merge(dt);
                currentPath++;
            }
            // updating primary pos of table2
            primary2_pos = exclude ? 1 : primary2_pos;

            // merged table dt1 and dt2
            DataTable dt3 = new DataTable();
            dt3.Merge(dt1, false, MissingSchemaAction.Add);
            dt3.Merge(dt2, false, MissingSchemaAction.Add);
            DataTable tmp = dt3.Clone();
            DataTable resultDt = tmp.Clone();

            // outer join two tables
            DataTable joinedDt = (from t1 in dt1.AsEnumerable()
                                  join t2 in dt2.AsEnumerable()
                                  on new { ID = t1[primary1_pos] } equals new { ID = t2[primary2_pos] }
                                  select tmp.LoadDataRow(Concatenate((object[])t1.ItemArray, (object[])t2.ItemArray, primary2_pos), true)).CopyToDataTable();

            if (joinedDt.Rows.Count != dt1.Rows.Count)
            {
                DataTable errorDt = dt1.AsEnumerable().Except(
                    from t1 in dt1.AsEnumerable() join t2 in dt2.AsEnumerable() on t1[primary1_pos] equals t2[primary2_pos] select t1).CopyToDataTable();

                Exception ex = new Exception("Some BOM records are not matching with CAD records.");
                ex.Data.Add("DATA_TABLE", (DataTable)errorDt);
                throw ex;
            }

            resultDt.Merge(joinedDt);
            // swapping column
            resultDt.Columns[3].SetOrdinal(1);
            resultDt.Columns[2].SetOrdinal(3);
            if (exclude == true)
            {
                resultDt.Columns.RemoveAt(3);
            }

            if (include != null && include.Length > 0)
            {
                bool available = dt2.AsEnumerable().Any(r => r[0].ToString() == include);
                if (available)
                {
                    DataRow r = resultDt.NewRow();
                    r[0] = null;
                    r[1] = include;
                    r[2] = "1";
                    resultDt.Rows.InsertAt(r, 0);
                    foreach (var (path, i) in paths1.Select((v, i) => (v, i)))
                    {
                        string filename = Path.GetFileNameWithoutExtension(path);
                        DataRow r2 = resultDt.NewRow();
                        r2[0] = include;
                        r2[1] = filename;
                        r2[2] = "1";
                        resultDt.Rows.InsertAt(r2, i + 1);
                    }
                }
                else
                {
                    MessageBox.Show(include + " not found.", "Warning");
                }
            }

            e.Result = resultDt;
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            lblResult.Text = (e.ProgressPercentage.ToString() + "%");
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            DateTime now = DateTime.Now;
            string outputPath = Properties.Settings.Default.output_dir + "\\" + now.ToString("ddMMyyyyHHmmss") + ".csv";
            string outputErrorPath = Properties.Settings.Default.output_dir + "\\" + "error_" + now.ToString("ddMMyyyyHHmmss") + ".csv";
            string delimiter = ((Delimiter)Properties.Settings.Default.output_delimiter).GetString();
            Encoding encoding = ((Encoder)Properties.Settings.Default.output_encoding).GetEncoding();

            if (e.Cancelled == true)
            {
                MessageBox.Show("Canceled");
                lblResult.Text = "Canceled!";
            }
            else if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Error");
                lblResult.Text = "Error: " + e.Error.Message;
                DataTable dt = (DataTable)e.Error.Data["DATA_TABLE"];
                try
                {
                    using (Stream s = File.Create(outputErrorPath))
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
                    MessageBox.Show("Error saved to the output directory.", "Finished");
                    var fullPath = System.IO.Path.GetFullPath(Properties.Settings.Default.output_dir);
                    Process.Start("explorer.exe", @"" + fullPath + "");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                progressBar1.Value = progressBar1.Maximum;
                lblResult.Text = "Done!";
                button1.Enabled = false;
                DataTable dt = (DataTable)e.Result;
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
                    this.usingFilename = true;
                    switch (type)
                    {
                        case 0:
                            this.delimiter = ((Delimiter)Properties.Settings.Default.cad_delimiter1).GetString();
                            this.hasHeaderRow = !!Properties.Settings.Default.cad_header1;
                            this.columnMatchPosition = (int)Properties.Settings.Default.cad_id_pos1;
                            this.columnQuantityPosition = (int)Properties.Settings.Default.cad_quantity_pos1;
                            break;
                        case 1:
                            this.delimiter = ((Delimiter)Properties.Settings.Default.cad_delimiter2).GetString();
                            this.hasHeaderRow = !!Properties.Settings.Default.cad_header2;
                            this.columnMatchPosition = (int)Properties.Settings.Default.cad_id_pos2;
                            this.columnQuantityPosition = (int)Properties.Settings.Default.cad_quantity_pos2;
                            break;
                        case 2:
                            this.delimiter = ((Delimiter)Properties.Settings.Default.cad_delimiter3).GetString();
                            this.hasHeaderRow = !!Properties.Settings.Default.cad_header3;
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

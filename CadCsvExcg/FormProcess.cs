using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Windows.Shapes;

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
                CSVConfig config = new CSVConfig((int)genericList[1]);
                DataTable dt = ReadFile(path, config.hasHeaderRow, config.delimiter, (complete) =>
                {
                    if (worker.CancellationPending == true)
                    {
                        e.Cancel = true;
                    }
                    else
                    {
                        worker.ReportProgress(complete);
                    }
                    return e.Cancel;
                });
                e.Result = dt;
            }
            catch (Exception ex)
            {
                throw ex;
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

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            List<object> genericList = e.Argument as List<object>;
            string[] pathList1 = (string[])genericList[0];
            string[] pathList2 = (string[])genericList[1];
            List<DataTable> dtList1 = new List<DataTable>();
            List<DataTable> dtList2 = new List<DataTable>();
            CSVConfig config1 = new CSVConfig(1);
            CSVConfig config2 = new CSVConfig(2);
            bool isExclude = !!Properties.Settings.Default.output_exclude;
            bool isInclude = !!Properties.Settings.Default.output_isinclude;
            string includeData = Properties.Settings.Default.output_include;

            bool isError = false;
            DataTable errorDt = new DataTable();

            // k = 0 ? 'CAD' : 'BOM'
            for (int k = 0; k < 2; k++)
            {
                List<DataTable> dtList = new List<DataTable>();
                string[] pathList = pathList1;
                CSVConfig config = config1;
                if (k == 1)
                {
                    pathList = pathList2;
                    config = config2;
                }

                foreach (var (path, i) in pathList.Select((v, i) => (v, i)))
                {
                    string filename = System.IO.Path.GetFileNameWithoutExtension(path);
                    DataTable dt = ReadFile(path, config.hasHeaderRow, config.delimiter, (complete) =>
                    {
                        if (worker.CancellationPending == true)
                        {
                            e.Cancel = true;
                        }
                        else
                        {
                            // TODO: add calculation of total percentage
                            worker.ReportProgress(complete);
                        }
                        return e.Cancel;
                    });
                    dt.TableName = filename;
                    dtList.Add(dt);
                }

                if (k == 0)
                {
                    dtList1 = dtList;
                }
                else
                {
                    dtList2 = dtList;
                }
            }

            /* --- WORKING CAD TABLES --- */

            foreach (DataTable dt in dtList1)
            {
                int match = config1.columnMatchPosition - 1;
                int quantity = config1.columnQuantityPosition - 1;

                // removing unnecessary columns
                for (int i = dt.Columns.Count - 1; i >= 0; i--)
                {
                    if (i != match && i != quantity)
                    {
                        dt.Columns.RemoveAt(i);
                    }
                }

                // renaming some columns
                dt.Columns[quantity].ColumnName = "員数";

                // removing duplicated rows and increasing quantity
                using (DataTable tmpDt = dt.AsEnumerable().GroupBy(r => new
                {
                    firstColumn = r[0],
                    match = r[match]
                }).Select(g =>
                {
                    var row = g.First();
                    row.SetField(dt.Columns[quantity].ColumnName.ToString(), g.Sum(r => Convert.ToInt32(r[quantity])).ToString());
                    return row;
                }).CopyToDataTable())
                {
                    dt.Clear();
                    dt.Merge(tmpDt, false, MissingSchemaAction.Add);
                }

                // prepending filename as column
                dt.Columns.Add("親品目番号").SetOrdinal(0);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i][0] = dt.TableName;
                }
            }

            // merging tables into one table if include option checked
            if (isInclude)
            {
                using (DataTable tmp = new DataTable())
                {
                    foreach (DataTable dt in dtList1)
                    {
                        tmp.Merge(dt);
                    }
                    dtList1.Clear();
                    dtList1.Add(tmp.DefaultView.ToTable(/*distinct*/ true));
                }
            }

            /* --- WORKING BOM TABLES --- */

            // merging tables into one table
            using (DataTable tmp = new DataTable())
            {
                foreach (DataTable dt in dtList2)
                {
                    dt.Columns[0].ColumnName = "品目番号";
                    tmp.Merge(dt);
                }
                dtList2.Clear();
                dtList2.Add(tmp.DefaultView.ToTable(/*distinct*/ true));
            }

            // excluding unneccessary columns if option checked
            if (isExclude)
            {
                int match = config2.columnMatchPosition - 1;
                for (int i = dtList2.First().Columns.Count - 1; i >= 0; i--)
                {
                    if (i == 0 || i == match) { }
                    else
                    {
                        dtList2.First().Columns.RemoveAt(i);
                    }
                }
            }

            /* --- JOINING TWO TABLES --- */

            foreach (var (dt, i) in dtList1.Select((v, i) => (v, i)))
            {
                DataTable dt1 = dt.Copy();
                DataTable dt2 = dtList2.First().Copy();
                DataTable resultDt = new DataTable();
                int match1, match2;
                string outputPath;
                // Detecting column positions to match
                if (config1.columnQuantityPosition > config1.columnMatchPosition)
                {
                    match1 = 1;
                }
                else
                {
                    match1 = 2;
                }
                if (isExclude)
                {
                    match2 = 1;
                }
                else
                {
                    match2 = config2.columnMatchPosition - 1;
                }

                DataTable mergedDt = new DataTable();
                mergedDt.Merge(dt1, false, MissingSchemaAction.Add);
                mergedDt.Merge(dt2, false, MissingSchemaAction.Add);
                using (DataTable tmp = mergedDt.Clone())
                {
                    DataTable joinedDt = (from t1 in dt1.AsEnumerable()
                                          join t2 in dt2.AsEnumerable()
                                          on new { ID = t1[match1] } equals new { ID = t2[match2] }
                                          select tmp.LoadDataRow(Concatenate((object[])t1.ItemArray, (object[])t2.ItemArray, match2), true)).CopyToDataTable();

                    var checkDt = dt1.AsEnumerable().Except(from t1 in dt1.AsEnumerable() join t2 in dt2.AsEnumerable() on t1[match1] equals t2[match2] select t1);
                    if (checkDt.Any())
                    {
                        isError = true;
                        MessageBox.Show("ファイル作成時にエラー発生しました。詳細はエラーファイルをご確認ください。");
                        // ADD TO ERROR DATATABLE
                    }


                    resultDt.Merge(joinedDt);
                }

                if (resultDt.Rows.Count != dt1.Rows.Count)
                {
                    // TODO: build error table
                    MessageBox.Show("Error" + resultDt.Rows.Count + "!=" + dt1.Rows.Count);
                    if (i == 0)
                    {
                        errorDt = dt1.Clone();
                    }
                    /*  DataTable tmp2 = dt1.AsEnumerable().Except(
                          from t1 in dt1.AsEnumerable() join t2 in dt2.AsEnumerable() on t1[match1] equals t2[match2] select t1).CopyToDataTable();

                          errorDt.Merge(tmp2);

                      errorDt.Columns.Add("エラーコード");
                      errorDt.Columns.Add("エラー内容");
                      errorDt.Rows[0][3] = "01";
                      errorDt.Rows[0][4] = "PLM品目にない構成品";
                      Exception ex = new Exception("ファイル作成時にエラー発生しました。詳細はエラーファイルをご確認ください。");
                      ex.Data.Add("DATA_TABLE", (DataTable)errorDt);
                      break;
                      throw ex;*/
                }
                else
                {
                }
                // managing column positions 
                resultDt.Columns[3].SetOrdinal(1);
                resultDt.Columns[2].SetOrdinal(3);
                if (isExclude)
                {
                    resultDt.Columns.RemoveAt(3);
                }

                if (isInclude && includeData != "")
                {
                    // adding additional rows
                    /*DataRow dr = resultDt.NewRow();
                    dr[0] = null;
                    dr[1] = includeData;
                    dr[2] = "1";
                    resultDt.Rows.InsertAt(dr, 0);*/

                    // show warning if includeData not found in tables;
                    if (!dt2.AsEnumerable().Any(r => r[0].ToString() == includeData))
                    {
                        MessageBox.Show(includeData + "がPLM品目マスタに存在しません。");
                    }

                    foreach (var (path, j) in pathList1.Select((v, j) => (v, j)))
                    {
                        string filename = System.IO.Path.GetFileNameWithoutExtension(path);
                    /*    DataRow r2 = resultDt.NewRow();
                        r2[0] = includeData;
                        r2[1] = filename;
                        r2[2] = "1";
                        resultDt.Rows.InsertAt(r2, j + 1);*/
                    }
                    outputPath = Properties.Settings.Default.output_dir + "\\" + includeData + "_out.csv";
                } else
                {
                    string filename = System.IO.Path.GetFileNameWithoutExtension(pathList1[i]);
                   /* DataRow dr = resultDt.NewRow();
                    dr[0] = null;
                    dr[1] = filename;
                    dr[2] = "1";
                    resultDt.Rows.InsertAt(dr, 0);*/
                    outputPath = Properties.Settings.Default.output_dir + "\\" + filename + "_out.csv";
                }
                SaveFile(resultDt, outputPath);
            }


            

            /*



                        List<DataTable> dts = new List<DataTable>();

                        // DataTable dt1 = new DataTable();
                        DataTable dt2 = new DataTable();

                        string[] paths1 = (string[])genericList[0];
                        string[] paths2 = (string[])genericList[1];


                        // CSVConfig outputConfig = new CSVConfig(3);

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
                            string filename = System.IO.Path.GetFileNameWithoutExtension(path);
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
                                        if (currentLine == 0 && config1.hasHeaderRow)
                                        {
                                            fields = PrependNewField(fields, (string)config1.columnFilenameText);
                                        }
                                        else
                                        {
                                            fields = PrependNewField(fields, (string)filename);
                                        }

                                        // setting header
                                        if (currentLine == 0)
                                        {
                                            foreach (var (field, i) in fields.Select((v, i) => (v, i + 1)))
                                            {
                                                string columnText = "桁";
                                                if (i == 1)
                                                {
                                                    columnText = "親品目番号";
                                                }
                                                else if (i == primary1_pos + 1)
                                                {
                                                    columnText = field;
                                                }
                                                else if (i == quantity_pos + 1)
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
                            if (currentPath == 0)
                            {
                                dts.Add(dt);
                            }
                            else
                            {
                                if (isInclude)
                                {
                                    dts.First().Merge(dt);
                                }
                                else
                                {
                                    dts.Add(dt);
                                }
                            }
                            currentPath++;
                        }
                        // BOM files
                        foreach (var path in (string[])genericList[1])
                        {
                            DataTable dt = new DataTable();
                            string filename = System.IO.Path.GetFileNameWithoutExtension(path);
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
                                                if (i == 1)
                                                {
                                                    columnText = "品目番号";
                                                }
                                                else if (i == primary2_pos + 1)
                                                {
                                                    columnText = field;
                                                }
                                                else
                                                {
                                                    if (config2.hasHeaderRow && !duplicated)
                                                    {
                                                        columnText = field;
                                                    }
                                                    else
                                                    {
                                                        columnText += (lastColumnIndex + i);
                                                    }
                                                }
                                                DataColumn dc = new DataColumn(columnText)
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

                        DateTime now = DateTime.Now;
                        Delimiter delimiter = (Delimiter)Properties.Settings.Default.output_delimiter;
                        Encoding encoding = ((Encoder)Properties.Settings.Default.output_encoding).GetEncoding();
                        int filecount = 0;

                        // Removing duplicated rows from dt2
                        DataTable tmpDt2 = dt2.DefaultView.ToTable( *//*distinct*//* true);
                        dt2.Clear();
                        dt2.Merge(tmpDt2, false, MissingSchemaAction.Add);

                        bool isError = false;

                        // merging two tables into one
                        foreach (var dt1 in dts)
                        {
                            // merging duplicated rows of CAD and increasing quantity
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

                            DataTable mergedDt = new DataTable();
                            mergedDt.Merge(dt1, false, MissingSchemaAction.Add);
                            mergedDt.Merge(dt2, false, MissingSchemaAction.Add);
                            DataTable tmp = mergedDt.Clone();
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

                                errorDt.Columns.Add("エラーコード");
                                errorDt.Columns.Add("エラー内容");
                                errorDt.Rows[0][3] = "01";
                                errorDt.Rows[0][4] = "PLM品目にない構成品";
                                Exception ex = new Exception("CADファイルに、PLM品目にない構成品が存在しています。");
                                ex.Data.Add("DATA_TABLE", (DataTable)errorDt);
                                throw ex;
                            }
                            else
                            {
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
                                    DataRow dr = resultDt.NewRow();
                                    dr[0] = null;
                                    dr[1] = include;
                                    dr[2] = "1";
                                    resultDt.Rows.InsertAt(dr, 0);

                                    // show warning if INCLUDED DATA not found;
                                    if (!dt2.AsEnumerable().Any(r => r[0].ToString() == include))
                                    {
                                        MessageBox.Show(include + "がPLM品目マスタに存在しません。", "Warning");
                                    }

                                    foreach (var (path, i) in paths1.Select((v, i) => (v, i)))
                                    {
                                        string filename = System.IO.Path.GetFileNameWithoutExtension(path);
                                        if (isInclude)
                                        {
                                            if (i == filecount)
                                            {
                                                DataRow r2 = resultDt.NewRow();
                                                r2[0] = include;
                                                r2[1] = filename;
                                                r2[2] = "1";
                                                resultDt.Rows.InsertAt(r2, 1);
                                            }
                                        }
                                        else
                                        {
                                            DataRow r2 = resultDt.NewRow();
                                            r2[0] = include;
                                            r2[1] = filename;
                                            r2[2] = "1";
                                            resultDt.Rows.InsertAt(r2, i + 1);
                                        }
                                    }
                                }
                                string outputPath = Properties.Settings.Default.output_dir + "\\" + include + "_out.csv";
                                if (dts.Count() > 1)
                                {
                                    string filename = System.IO.Path.GetFileNameWithoutExtension(paths1[filecount]);
                                    outputPath = Properties.Settings.Default.output_dir + "\\" + filename + "_out.csv";
                                    filecount++;
                                }
                                SaveFile(resultDt, outputPath, encoding, delimiter);
                            }
                        }
                        e.Result = "success";*/
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            lblResult.Text = (e.ProgressPercentage.ToString() + "%");
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            DateTime now = DateTime.Now;
            string outputErrorPath = Properties.Settings.Default.output_dir + "\\" + "error(" + now.ToString("yyyyMMddHHmmss") + ").csv";
            Delimiter delimiter = (Delimiter)Properties.Settings.Default.output_delimiter;
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
                SaveFile(dt, outputErrorPath);
            }
            else
            {
                progressBar1.Value = progressBar1.Maximum;
                lblResult.Text = "ファイル作成成功しました。";
                button1.Enabled = false;
                try
                {
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

        private static string[] PrependNewField(string[] target, string data)
        {
            List<string> list = new List<string>(target);
            list.Insert(0, data);
            return list.ToArray();
        }

        public static T[] Concatenate<T>(T[] array1, T[] array2, int excludedIndex = 0)
        {
            T[] result = new T[array1.Length + array2.Length - 1];
            array1.CopyTo(result, 0);
            array2.Where((v, i) => i != excludedIndex).ToArray().CopyTo(result, array1.Length);
            return result;
        }

        private static DataTable ReadFile(string path, bool isHeader, string delimiter, Func<int, bool> fn)
        {
            DataTable dt = new DataTable();
            long totalLines = File.ReadLines(path).Count();
            using (TextFieldParser reader = new TextFieldParser(path))
            {
                reader.TextFieldType = FieldType.Delimited;
                reader.SetDelimiters(new string[] { delimiter });
                reader.HasFieldsEnclosedInQuotes = true;
                long currentLine = 0;
                int totalComplete = 0;
                while (!reader.EndOfData)
                {
                    string[] fields = reader.ReadFields();
                    for (int i = 0; i < fields.Length; i++)
                    {
                        if (fields[i] == "")
                        {
                            fields[i] = null;
                        }
                    }
                    if (currentLine == 0)
                    {
                        bool isDuplicatedColumnName = fields.Length != fields.Distinct().Count();
                        foreach (var (field, i) in fields.Select((v, i) => (v, i)))
                        {
                            string columnName = "column";
                            if (!isHeader)
                            {
                                columnName += "_" + i;
                            }
                            else if (isDuplicatedColumnName)
                            {
                                columnName += "_" + i;
                            }
                            else
                            {
                                columnName = field;
                            }

                            DataColumn dc = new DataColumn(columnName)
                            {
                                AllowDBNull = true
                            };
                            dt.Columns.Add(dc);
                        }
                        if (!isHeader || isDuplicatedColumnName)
                        {
                            dt.Rows.Add(fields);
                        }
                    }
                    else
                    {
                        dt.Rows.Add(fields);
                    }
                    currentLine = reader.LineNumber;
                    int currentComplete = (int)Math.Round((double)(100 * currentLine / totalLines));
                    if (currentComplete > totalComplete)
                    {
                        totalComplete = currentComplete;
                        if (fn(totalComplete))
                        {
                            break;
                        }
                    }
                }
            }
            return dt;
        }

        private void SaveFile(DataTable dt, string dir)
        {
            Delimiter delimiter = (Delimiter)Properties.Settings.Default.output_delimiter;
            Encoding encoding = ((Encoder)Properties.Settings.Default.output_encoding).GetEncoding();
            if (File.Exists(dir))
            {
                string filename = System.IO.Path.GetFileName(dir);
                if (MessageBox.Show(filename + " は既に存在しています。上書きしてもよろしいでしょうか", "", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
            }
            using (Stream s = File.Create(dir))
            {
                StreamWriter sw = new StreamWriter(s, encoding);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sw.Write(dt.Columns[i]);
                    if (i < dt.Columns.Count - 1)
                    {
                        sw.Write(delimiter.GetString());
                    }
                }
                sw.Write(sw.NewLine);
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
                            sw.Write(delimiter.GetString());
                        }
                    }
                    sw.Write(sw.NewLine);
                }
                sw.Close();
            }
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
                case 3:
                    // output config
                    break;
            }
        }
    }
}
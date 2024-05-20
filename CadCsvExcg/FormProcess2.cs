using CsvHelper.Configuration.Attributes;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using CsvHelper;
using System.IO;
using System.Dynamic;
using System.Windows.Controls;

namespace CadCsvExcg
{
    public partial class FormProcess2 : Form
    {
        private BackgroundWorker _worker1, _worker2, _worker3;

        private List<string> _fileList1, _fileList2;

        private int _px1, _qx, _px2, _sx;

        private Delimiter _delimiter1, _delimiter2, _outDelimiter;
        private bool _hasHeader1, _hasHeader2, _isExclude, _isInclude, _isReady1 = false, _isReady2 = false;

        private string _includeName, _outPath;

        private Encoding _outEncoding;

        private List<Cad> _records1;
        private List<Bom> _records2;

        private FormProgress _progressForm;

        private int _complete = 0;

        public FormProcess2(IFormProcessData formData)
        {
            InitializeComponent();

            _fileList1 = formData.FileList1;
            _fileList2 = formData.FileList2;
            _px1 = formData.Px1;
            _px2 = formData.Px2;
            _qx = formData.Qx;
            _sx = formData.Sx;
            _delimiter1 = formData.Delimiter1;
            _delimiter2 = formData.Delimiter2;
            _outDelimiter = formData.Delimiter3;
            _hasHeader1 = formData.HasHeader1;
            _hasHeader2 = formData.HasHeader2;
            _isExclude = formData.IsExclude;
            _isInclude = formData.IsInclude;
            _includeName = formData.IncludeName;
            _outPath = formData.Output;
            _outEncoding = formData.Encode;

            _progressForm = new FormProgress();
            _progressForm.CancelEvent += ProgressForm_CancelEvent;


            _worker1 = new BackgroundWorker();
            _worker1.WorkerReportsProgress = true;
            _worker1.WorkerSupportsCancellation = true;
            _worker1.DoWork += ReadCad;
            _worker1.RunWorkerCompleted += ReadCadCompleted;
            _worker1.ProgressChanged += ReportProgress;

            _worker2 = new BackgroundWorker();
            _worker2.WorkerReportsProgress = true;
            _worker2.WorkerSupportsCancellation = true;
            _worker2.DoWork += ReadBom;
            _worker2.RunWorkerCompleted += ReadBomCompleted;
            _worker2.ProgressChanged += ReportProgress;

            _worker3 = new BackgroundWorker();
            _worker3.WorkerReportsProgress = true;
            _worker3.WorkerSupportsCancellation = true;
            _worker3.DoWork += Combine;
            _worker3.RunWorkerCompleted += CombineCompleted;
            _worker3.ProgressChanged += ReportProgress;

        }

        private void ProgressForm_CancelEvent(object sender, EventArgs e)
        {
            //Terminate();
        }

        private void FormProcess2_Load(object sender, EventArgs e)
        {
            StartRead();
            _progressForm.ShowDialog();
        }

        private void StartRead()
        {
            if (!_worker1.IsBusy && !_worker2.IsBusy)
            {
                _isReady1 = false;
                _isReady2 = false;
                _worker1.RunWorkerAsync();
                _worker2.RunWorkerAsync();
            }
        }

        private void StartCombine()
        {
            if (_isReady1 && _isReady2 && !_worker3.IsBusy)
            {
                _worker3.RunWorkerAsync();
            }
        }

        private void ReadCad(object sender, DoWorkEventArgs e)
        {
            try
            {
                BackgroundWorker worker = sender as BackgroundWorker;

                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = _hasHeader1,
                    Delimiter = _delimiter1.GetString(),
                    MissingFieldFound = null,
                };

                List<Cad> allRecords = new List<Cad>();

                int currentReading = 0;
                int totalFilesCount = _fileList1.Count;

                foreach (string file in _fileList1)
                {
                    currentReading++;
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                    }
                    else
                    {
                        var currentComplete = (int)Math.Round((double)((currentReading * 40) / totalFilesCount));
                        _complete += currentComplete;
                        worker.ReportProgress(_complete);
                    }

                    string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                    using (var reader = new StreamReader(file))
                    using (var csv = new CsvReader(reader, config))
                    {
                        csv.Context.RegisterClassMap(new CadMap(fileName, _px1, _qx, _sx));
                        var records = csv.GetRecords<Cad>();
                        var summarizedRecords = records
                            .GroupBy(r => new { r.FileName, r.PartNo })
                            .Select(g => new Cad
                            {
                                FileName = g.Key.FileName,
                                PartNo = g.Key.PartNo,
                                Quantity = g.Sum(r => r.Quantity)
                            });
                        allRecords.AddRange(summarizedRecords.ToList());
                    }

                }
                e.Result = allRecords
                    .OrderBy(record => record.FileName)
                    .ThenBy(record => record.PartNo)
                    .ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ReadCadCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show($"An error occurred when reading cad file: {e.Error.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Terminate();
            }
            else if (e.Cancelled)
            {
                MessageBox.Show("Process cancelled when reading cad file", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Terminate();
            }
            else
            {
                _records1 = (List<Cad>)e.Result;
                _isReady1 = true;
                StartCombine();
            }
        }

        private void ReadBom(object sender, DoWorkEventArgs e)
        {
            try
            {
                BackgroundWorker worker = sender as BackgroundWorker;

                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = _hasHeader2,
                    Delimiter = _delimiter2.GetString(),
                    MissingFieldFound = null,
                };

                List<Bom> allRecords = new List<Bom>();

                int currentReading = 0;
                int totalFilesCount = _fileList1.Count;

                foreach (string file in _fileList2)
                {
                    currentReading++;
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                    }
                    else
                    {
                        var currentComplete = (int)Math.Round((double)((currentReading * 40) / totalFilesCount));
                        _complete += currentComplete;
                        worker.ReportProgress(_complete);
                    }
                    using (var reader = new StreamReader(file))
                    using (var csv = new CsvReader(reader, config))
                    {
                        csv.Context.RegisterClassMap(new BomMap(_px2, _sx));
                        var records = csv.GetRecords<Bom>();
                        allRecords.AddRange(records.ToList());
                    }
                }

                // Removed duplicated rows
                var filteredAllRecords = allRecords
                    .GroupBy(r => new
                    {
                        itemNo = r.ItemNo,
                        partNo = r.PartNo
                    })
                    .Select(g => g.First())
                    .ToList();

                e.Result = filteredAllRecords
                    .OrderBy(record => record.ItemNo)
                    .ThenBy(record => record.PartNo)
                    .ToList();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ReadBomCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show($"An error occurred when reading bom file: {e.Error.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Terminate();
            }
            else if (e.Cancelled)
            {
                MessageBox.Show("Process cancelled when reading bom file", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Terminate();
            }
            else
            {
                _records2 = (List<Bom>)e.Result;
                foreach (var record in _records2)
                _isReady2 = true;
                StartCombine();
            }
        }

        private void Combine(object sender, DoWorkEventArgs e)
        {
            try
            {
                BackgroundWorker worker = sender as BackgroundWorker;

                bool isError = false;

                var errorRecords = new List<ErrorOutput>();

                var joinedRecords = from cad in _records1
                                    join bom in _records2
                                    on cad.PartNo equals bom.PartNo
                                    where bom.IsSkip != true
                                    select new
                                    {
                                        FileName = cad.FileName,
                                        PartNo = cad.PartNo,
                                        Quantity = cad.Quantity,
                                        ItemNo = bom.ItemNo,
                                        OtherColumns = bom.OtherColumns,
                                    };

                var skippedRecords = from cad in _records1
                                    join bom in _records2
                                    on cad.PartNo equals bom.PartNo
                                    where bom.IsSkip
                                    select new
                                    {
                                        FileName = cad.FileName,
                                        PartNo = cad.PartNo,
                                        Quantity = cad.Quantity,
                                        ItemNo = bom.ItemNo,
                                        OtherColumns = bom.OtherColumns,
                                    };

                var notJoinedRecords = from cad in _records1
                                       join bom in _records2
                                       on cad.PartNo equals bom.PartNo into joined
                                       from bom in joined.DefaultIfEmpty()
                                       where bom == null
                                       select new
                                       {
                                           FileName = cad.FileName,
                                           PartNo = cad.PartNo,
                                           Quantity = cad.Quantity,
                                           ItemNo = bom?.ItemNo,
                                           OtherColumns = bom?.OtherColumns
                                       };

                var duplicatedRecords = from j in joinedRecords
                                        group j by new { j.FileName, j.PartNo } into g
                                        where g.Count() > 1
                                        from row in g
                                        select row;

                if (skippedRecords.Any())
                {
                    isError = true;
                    errorRecords.AddRange(
                        skippedRecords.Select(r => new ErrorOutput
                        {
                            FileName = r.FileName,
                            PartNo = r.PartNo,
                            Quantity = r.Quantity,
                            ErrorCode = "03",
                            ErrorDesc = "使用禁止部品です。代替品を検討してください。"
                        })
                    );
                }

                if (notJoinedRecords.Any())
                {
                    isError = true;
                    errorRecords.AddRange(
                        notJoinedRecords.Select(r => new ErrorOutput
                        {
                            FileName = r.FileName,
                            PartNo = r.PartNo,
                            Quantity = r.Quantity,
                            ErrorCode = "01",
                            ErrorDesc = "PLM品目にない構成品。",
                        })
                    );
                }

                if (duplicatedRecords.Any())
                {
                    isError = true;
                    errorRecords.AddRange(
                        duplicatedRecords.Select(r => new ErrorOutput
                        {
                            FileName = r.FileName,
                            PartNo = r.PartNo,
                            Quantity = r.Quantity,
                            ErrorCode = "02",
                            ErrorDesc = "構成品がPLM品目マスタに重複されています。",
                        })
                    );
                }

                var outputConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = true,
                    Delimiter = _outDelimiter.GetString(),
                    MissingFieldFound = null,
                };

                var otherColumnsWithEmptyValue = new Dictionary<string, string>();
                foreach (var kvp in joinedRecords.First().OtherColumns)
                {
                    otherColumnsWithEmptyValue.Add(kvp.Key, "");
                }

                if (isError)
                {
                    DateTime now = DateTime.Now;
                    string outputFileName = $"error_{now.ToString("yyyyMMddHHmmss")}";

                    using (var writer = new StreamWriter($"{_outPath}\\{outputFileName}.csv", false, _outEncoding))
                    using (var csv = new CsvWriter(writer, outputConfig))
                    {
                        csv.Context.RegisterClassMap<ErrorOutputMap>();
                        csv.WriteRecords(errorRecords);
                    }
                }
                
                if (_isInclude && !String.IsNullOrEmpty(_includeName))
                {
                    var records = joinedRecords.Select(r => (dynamic)r).ToList();

                    int index = 0;
                    foreach (var group in joinedRecords.GroupBy(r => r.FileName))
                    {
                        // TODO: shift to left and add new column name with includeName
                        dynamic newRecord = new ExpandoObject();
                        newRecord.FileName = $"CT-{_includeName}";
                        newRecord.PartNo = group.Key;
                        newRecord.Quantity = 1;
                        newRecord.ItemNo = "";
                        newRecord.OtherColumns = otherColumnsWithEmptyValue;

                        records.Insert(index, newRecord);
                        index++;
                    }

                    dynamic prependRecord = new ExpandoObject();
                    prependRecord.FileName = "";
                    prependRecord.PartNo = $"CT-{_includeName}";
                    prependRecord.Quantity = 1;
                    prependRecord.ItemNo = "";
                    prependRecord.OtherColumns = otherColumnsWithEmptyValue;
                    records.Insert(0, prependRecord);

                    var resultRecords = new List<dynamic>();
                    var isIncludeNameInData = false;
                    foreach (var record in records)
                    {
                        dynamic expando = new ExpandoObject();
                        ((IDictionary<string, object>)expando)["親品目番号"] = record.FileName;
                        ((IDictionary<string, object>)expando)["子品目番号"] = record.PartNo;
                        ((IDictionary<string, object>)expando)["員数"] = record.Quantity;
                        if (!_isExclude && record.OtherColumns != null && record.OtherColumns is Dictionary<string, string>)
                        {
                            ((IDictionary<string, object>)expando)["品目番号"] = record.ItemNo;
                            foreach (var kvp in record.OtherColumns)
                            {
                                string key = kvp.Key.ToString();
                                string value = kvp.Value.ToString();
                                ((IDictionary<string, object>)expando)[key] = value;
                            }
                        }

                        if (record.FileName == $"CT-{_includeName}")
                        {
                            isIncludeNameInData = true;
                        }
                        resultRecords.Add(expando);
                    }

                    if (!isIncludeNameInData)
                    {
                        MessageBox.Show($"CT-\"{_includeName}\"がPLM品目マスタに存在しません。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    using (var writer = new StreamWriter($"{_outPath}\\CT-{_includeName}_out.csv", false, _outEncoding))
                    using (var csv = new CsvWriter(writer, outputConfig))
                    {
                        csv.WriteRecords(resultRecords);
                    }
                }
                else
                {
                    foreach (var group in joinedRecords.GroupBy(r => r.FileName))
                    {
                        string fileName = group.Key;
                        var records = group.Select(r => (dynamic)r).ToList();
                        dynamic prependRecord = new ExpandoObject();
                        prependRecord.FileName = "";
                        prependRecord.PartNo = fileName;
                        prependRecord.Quantity = 1;
                        prependRecord.ItemNo = "";
                        prependRecord.OtherColumns = otherColumnsWithEmptyValue;
                        records.Insert(0, prependRecord);

                        var resultRecords = new List<dynamic>();
                        foreach (var record in records)
                        {
                            dynamic expando = new ExpandoObject();
                            ((IDictionary<string, object>)expando)["親品目番号"] = record.FileName;
                            ((IDictionary<string, object>)expando)["子品目番号"] = record.PartNo;
                            ((IDictionary<string, object>)expando)["員数"] = record.Quantity;
                            if (!_isExclude && record.OtherColumns != null && record.OtherColumns is Dictionary<string, string>)
                            {
                                ((IDictionary<string, object>)expando)["品目番号"] = record.ItemNo;
                                foreach (var kvp in record.OtherColumns)
                                {
                                    string key = kvp.Key.ToString();
                                    string value = kvp.Value.ToString();
                                    ((IDictionary<string, object>)expando)[key] = value;
                                }
                            }
                            resultRecords.Add(expando);
                        }

                        using (var writer = new StreamWriter($"{_outPath}\\{fileName}_out.csv", false, _outEncoding))
                        using (var csv = new CsvWriter(writer, outputConfig))
                        {
                            csv.WriteRecords(resultRecords);
                        }
                    }
                }

                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                }
                else
                {   
                    _complete += 20;
                    worker.ReportProgress(_complete);
                }

                e.Result = !isError;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void CombineCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                MessageBox.Show($"An error occurred when combining: {e.Error.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (e.Cancelled)
            {
                MessageBox.Show("Process cancelled when combining", "Cancelled", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                var isSuccess = (bool)e.Result;

                if (isSuccess)
                {
                    MessageBox.Show("Done", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("ファイル作成時にエラー発生しました。詳細はエラーファイルをご確認ください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            Terminate();
        }

        private void ReportProgress(object sender, ProgressChangedEventArgs e)
        {
            _progressForm.UpdateProgressBar(e.ProgressPercentage);
        }

        private void Terminate()
        {
            if (_worker1.WorkerSupportsCancellation == true)
            {
                _worker1.CancelAsync();
            }
            if (_worker2.WorkerSupportsCancellation == true)
            {
                _worker2.CancelAsync();
            }
            if (_worker3.WorkerSupportsCancellation == true)
            {
                _worker3.CancelAsync();
            }
            _progressForm.CancelEvent -= ProgressForm_CancelEvent;
            _progressForm.Close();
            this.Close();
        }
    }
    public class Cad
    {
        public string FileName { get; set; }
        public string PartNo { get; set; }
        public int Quantity { get; set; }
    }
    public sealed class CadMap : ClassMap<Cad>
    {
        public CadMap(string fileName, int px, int qx, int sx)
        {
            Map(m => m.FileName).Constant(fileName);
            Map(m => m.PartNo).Index(px);
            Map(m => m.Quantity).Convert(args =>
            {
                string quantityStr = args.Row.GetField<string>(qx);
                if (int.TryParse(quantityStr, out int quantity))
                {
                    return quantity;
                }
                return 0;
            });
        }
    }
    public class Bom
    {
        public string ItemNo { get; set; }
        public string PartNo { get; set; }
        public bool IsSkip { get; set; }

        [Optional]
        public Dictionary<string, string> OtherColumns { get; set; } = new Dictionary<string, string>();
    }
    public sealed class BomMap : ClassMap<Bom>
    {
        public BomMap(int px, int sx)
        {
            Map(m => m.ItemNo).Index(0);
            Map(m => m.PartNo).Index(px);
            Map(m => m.IsSkip).Convert(args =>
            {
                string skipCol = args.Row.GetField<string>(sx);
                if (skipCol == "使用不可")
                {
                    return true;
                }
                return false;
            });
            Map(m => m.OtherColumns).Convert(args =>
            {
                var fieldHeaders = args.Row.HeaderRecord;
                var otherColumns = new Dictionary<string, string>();
                int columnIndex = 2;
                for (int i = 1; i < args.Row.ColumnCount; i++)
                {
                    if (i != px && i != sx)
                    {
                        var columnName = fieldHeaders != null ? fieldHeaders[i] : $"Column{columnIndex}";
                        var columnValue = args.Row.GetField<string>(i);
                        if (!string.IsNullOrWhiteSpace(columnValue))
                        {
                            otherColumns[columnName] = columnValue;
                        }
                        columnIndex++;
                    }
                }
                return otherColumns;
            });

        }
    }

    public class ErrorOutput : Cad
    {
        public string ErrorCode { get; set; }
        public string ErrorDesc { get; set; }
    }

    public class ErrorOutputMap : ClassMap<ErrorOutput>
    {
        public ErrorOutputMap()
        {
            Map(m => m.FileName).Name("親品目番号");
            Map(m => m.PartNo).Name("子品目番号");
            Map(m => m.Quantity).Name("員数");
            Map(m => m.ErrorCode).Name("エラーコード");
            Map(m => m.ErrorDesc).Name("エラー内容");
        }
    }

    public interface IFormProcessData
    {
        List<string> FileList1 { get; }
        List<string> FileList2 { get; }
        Delimiter Delimiter1 { get; }
        Delimiter Delimiter2 { get; }
        Delimiter Delimiter3 { get; }
        int Px1 { get; }
        int Px2 { get; }
        int Qx { get; }
        int Sx { get; }
        bool HasHeader1 { get; }
        bool HasHeader2 { get; }
        bool IsExclude { get; }
        bool IsInclude { get; }
        string IncludeName { get; }
        string Output { get; }
        Encoding Encode { get; }
    }

    public class FormProcessData : IFormProcessData
    {
        public FormProcessData(
            List<string> fileList1,
            List<string> fileList2,
            Delimiter delimiter1,
            Delimiter delimiter2,
            Delimiter delimiter3,
            int px1,
            int px2,
            int qx,
            int sx,
            bool hasHeader1,
            bool hasHeader2,
            bool isExlcude,
            bool isInclude,
            string includeName,
            string output,
            Encoding encode
            )
        {
            FileList1 = fileList1;
            FileList2 = fileList2;
            Delimiter1 = delimiter1;
            Delimiter2 = delimiter2;
            Delimiter3 = delimiter3;
            Px1 = px1;
            Px2 = px2;
            Qx = qx;
            Sx = sx;
            HasHeader1 = hasHeader1;
            HasHeader2 = hasHeader2;
            IsExclude = isExlcude;
            IsInclude = isInclude;
            IncludeName = includeName;
            Output = output;
            Encode = encode;
        }

        public List<string> FileList1 { get; }
        public List<string> FileList2 { get; }
        public Delimiter Delimiter1 { get; }
        public Delimiter Delimiter2 { get; }
        public Delimiter Delimiter3 { get; }
        public int Px1 { get; }
        public int Px2 { get; }
        public int Qx { get; }
        public int Sx { get; }
        public bool HasHeader1 { get; }
        public bool HasHeader2 { get; }
        public bool IsExclude { get; }
        public bool IsInclude { get; }
        public string IncludeName { get; }
        public string Output { get; }
        public Encoding Encode { get; }
    }
}

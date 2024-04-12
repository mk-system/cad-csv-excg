using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace CadCsvExcg {
    public static class Utils
    {
        public static bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                {
                    return false;
                }
            }
            return true;
        }

        public static DataTable readCsv(string filePath, Delimiter delimiter, bool hasHeader, string headerPrefix = "column")
        {
            DataTable dt = new DataTable();
            using (TextFieldParser reader = new TextFieldParser(filePath))
            {
                long totalLines = File.ReadLines(filePath).Count();
                reader.TextFieldType = FieldType.Delimited;
                reader.SetDelimiters(new string[] { delimiter.GetString() });
                reader.HasFieldsEnclosedInQuotes = true;
                long currentLine = 0;
                while (!reader.EndOfData)
                {
                    string[] fields = reader.ReadFields();

                    if (currentLine == 0)
                    {
                        foreach (var (field, i) in fields.Select((v, i) => (v, i)))
                        {
                            string caption;
                            string columnName = headerPrefix + i;

                            if (hasHeader)
                            {
                                caption = field.ToString();
                            }
                            else
                            {
                                caption = columnName;
                            }

                            DataColumn dc = new DataColumn
                            {
                                ColumnName = columnName,
                                // DataType = typeof(string),
                                Caption = caption,
                            };

                            dt.Columns.Add(dc);
                        }

                        if (!hasHeader)
                        {
                            dt.Rows.Add(fields);
                        }
                    }
                    else
                    {
                        dt.Rows.Add(fields);
                    }

                    if (reader.LineNumber == -1)
                    {
                        break;
                    }
                    currentLine = reader.LineNumber;
                }
            }
            return dt;
        }

        public static void saveCsv(DataTable dataTable, string filePath, Delimiter delimiter, Encoding encoding)
        {
            if (File.Exists(filePath))
            {
                string filename = System.IO.Path.GetFileName(filePath);
                if (MessageBox.Show(filename + " は既に存在しています。上書きしてもよろしいでしょうか", "", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
            }
            using (Stream stream = File.Create(filePath))
            {
                StreamWriter writer = new StreamWriter(stream, encoding);
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    writer.Write(dataTable.Columns[i]);
                    if (i < dataTable.Columns.Count - 1)
                    {
                        writer.Write(delimiter.GetString());
                    }
                }
                writer.Write(writer.NewLine);
                foreach (DataRow row in dataTable.Rows)
                {
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        if (!Convert.IsDBNull(row[i]))
                        {
                            string value = row[i].ToString();
                            if (value.Contains(","))
                            {
                                value = String.Format("\"{0}\"", value);
                            }
                            writer.Write(value);
                        }
                        if (i < dataTable.Columns.Count - 1)
                        {
                            writer.Write(delimiter.GetString());
                        }
                    }
                    writer.Write(writer.NewLine);
                }
                writer.Close();
            }
        }
    }
}

using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace CadCsvExcg
{
    public static class CSVUtility
    {
        public static DataTable Import(string csv_file_path, string delimiter, int colStart = 1, bool header = false, int primary = 0, bool append = false, string fileColumnName = "FILENAME")
        {
            DataTable csvData = new DataTable();
            DataColumn[] keys = new DataColumn[1];
            using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
            {
                csvReader.TextFieldType = FieldType.Delimited;
                csvReader.SetDelimiters(new string[] { delimiter });
                csvReader.HasFieldsEnclosedInQuotes = true;

                // First line of file
                string[] colFields = csvReader.ReadFields();

                if (append)
                {
                    if (header)
                    {
                        colFields = InsertData(colFields, fileColumnName);
                    } else
                    {
                        colFields = InsertData(colFields, Path.GetFileNameWithoutExtension(csv_file_path));
                    }
                }

                foreach (var (column, i) in colFields.Select((v, i) => (v, i)))
                {
                    DataColumn dc = new DataColumn(header ? column : ("COLUMN " + (colStart + i + 1)));
                    dc.AllowDBNull = true;
                    
                    csvData.Columns.Add(dc);
                    if (primary > 0 && primary - 1 == i)
                    {
                        keys[0] = dc;
                    }
                }

                if (!header)
                {
                    csvData.Rows.Add(colFields);
                }

                while (!csvReader.EndOfData)
                {
                    string[] fieldData = csvReader.ReadFields();
                    if (append)
                    {
                        fieldData = InsertData(fieldData, Path.GetFileNameWithoutExtension(csv_file_path));
                    }
                    //Making empty value as null
                    for (int i = 0; i < fieldData.Length; i++)
                    {
                        if (fieldData[i] == "")
                        {
                            fieldData[i] = null;
                        }
                    }
                    csvData.Rows.Add(fieldData);
                }
            }
            if (primary > 0)
            {
                csvData.PrimaryKey = keys;
            }
            return csvData;
        }

        public static void Export(DataTable dt, string outputPath, string delimiter = ",", bool header = true)
        {
            using (Stream s = File.Create(outputPath))
            {
                StreamWriter sw = new StreamWriter(s);
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
        }

        private static string[] InsertData(string[] target, string data, int index = 0)
        {
            List<string> list = new List<string>(target);
            list.Insert(index, data);
            return list.ToArray();
        }
    }
}

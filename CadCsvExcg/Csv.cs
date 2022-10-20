using Microsoft.VisualBasic.FileIO;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace CadCsvExcg
{
    class Csv
    {
        public static DataTable Parse(string csv_file_path, string delimiter, bool header = false, int primary = 0, bool append = false)
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
                        colFields = InsertData(colFields, "FILENAME");
                    } else
                    {
                        colFields = InsertData(colFields, Path.GetFileNameWithoutExtension(csv_file_path));
                    }
                }

                foreach (var (column, i) in colFields.Select((v, i) => (v, i)))
                {
                    DataColumn dc = new DataColumn(header ? column : ("COLUMN " + (i + 1)));
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

        private static string[] InsertData(string[] target, string data, int index = 0)
        {
            List<string> list = new List<string>(target);
            list.Insert(index, data);
            return list.ToArray();
        }
    }
}

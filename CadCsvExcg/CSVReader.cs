using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CadCsvExcg
{
    public class CSVReader
    {

        public DataSet ReadCSVFile(string fullPath, bool headerRow)
        {
            string path = Path.GetDirectoryName(fullPath);
            string filename = Path.GetFileName(fullPath);

            DataSet ds = new DataSet();

            try
            {
                if (File.Exists(fullPath))
                {
                    string ConStr = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}" + ";Extended Properties=\"Text;characterset=65001;HDR={1};FMT=Delimited\"", path, headerRow ? "Yes" : "No");
                    string SQL = string.Format("SELECT * FROM {0}", filename);
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(SQL, ConStr))
                    {
                        adapter.Fill(ds);
                        ds.Tables[0].TableName = "Table1";
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return ds;
        }
    }
}

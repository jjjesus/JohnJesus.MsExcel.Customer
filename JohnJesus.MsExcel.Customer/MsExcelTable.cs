using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JohnJesus.MsExcel.Customer
{
    internal class MsExcelTable
    {
        public MsExcelTable() { }

        public static DataSet ImportExcelXLS(string fileName)
        {
            DataSet output = new DataSet();

            try
            {
                string connName = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
                    fileName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';";
                using (var conn = new OleDbConnection(connName))
                {
                    conn.Open();
                    var dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                        new object[] { null, null, null, "TABLE" });

                    foreach (DataRow row in dt.Rows)
                    {
                        string sheet = row["TABLE_NAME"].ToString();
                        var cmd = new OleDbCommand(
                            "SELECT * FROM [" + sheet + "]", conn);
                        cmd.CommandType = CommandType.Text;

                        var outputTable = new DataTable(sheet);
                        output.Tables.Add(outputTable);
                        new OleDbDataAdapter(cmd).Fill(outputTable);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return output;
        }
    }
}

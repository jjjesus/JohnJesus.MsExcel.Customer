using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace JohnJesus.MsExcel.Customer
{
    public class MsExcelParser
    {
        private const string MS_EXCEL_FILE_NAME = "Customers.xls";
        private const string EMBEDDED_NAMESPACE = "JohnJesus.MsExcel.Customer.Data";

        public MsExcelParser() { }

        public void Parse()
        {
            string filePath = createMsExcelFile();
            DataSet dataSet = MsExcelTable.ImportExcelXLS(filePath);
            var data = dataSet.Tables[0].AsEnumerable();
            parseData(data);
        }

        private void parseData(EnumerableRowCollection<DataRow> data)
        {
            var results = from _customer in data
                          select new Customer
                          {
                              Id = _customer.Field<double>("Id").ToString(),
                              Name = _customer.Field<string>("Name"),
                              SalesAmount = _customer.Field<double>("Sales_Amount").ToString()
                          };
            foreach (var customer in results)
            {
                Console.WriteLine("{0}:{1}:{2}", customer.Id, customer.Name, customer.SalesAmount);
            }
        }

        private string createMsExcelFile()
        {
            string outputDir = Path.GetTempPath();
            string outputPath = string.Format("{0}{1}", outputDir, MS_EXCEL_FILE_NAME);
            string embeddedMsExcelFile = string.Format("{0}.{1}", EMBEDDED_NAMESPACE, MS_EXCEL_FILE_NAME);
            var asm = Assembly.GetExecutingAssembly();
            List<string> resourceNames = asm.GetManifestResourceNames().ToList();
            using (var input = asm.GetManifestResourceStream(embeddedMsExcelFile))
            {
                using (FileStream output = File.Open(outputPath, FileMode.Create))
                {
                    copyStream(input, output);
                }
            }
            return outputPath;
        }

        private void copyStream(Stream input, FileStream output)
        {
            byte[] buffer = new byte[8 * 1024];
            int len = 0;
            while ((len = input.Read(buffer, 0, buffer.Length)) > 0)
            {
                output.Write(buffer, 0, len);
            }
        }
    }
}

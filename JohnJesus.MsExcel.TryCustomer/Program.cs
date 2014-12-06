using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using JohnJesus.MsExcel.Customer;

namespace JohnJesus.MsExcel.TryCustomer
{
    class Program
    {
        static void Main(string[] args)
        {
            var msExcelParser = new MsExcelParser();
            msExcelParser.Parse();
            Console.WriteLine("Hit ENTER to continue");
            Console.ReadLine();
        }
    }
}

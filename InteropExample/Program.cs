
using InteropExample.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropExample
{
    class Program
    {
        static void Main(string[] args)
        {
            OfflineExcel.CreateExcelFile(@"E:\test.xlsx", "test");
            OfflineExcel.InsertTextExistingExcel(@"E:\test.xlsx", "test 1", "A", 1);
            OfflineExcel.InsertTextExistingExcel(@"E:\test.xlsx", "test 2", "A", 4);
            OfflineExcel.InsertTextExistingExcel(@"E:\test.xlsx", "test 3", "B", 4);
            OfflineExcel.InsertTextExistingExcel(@"E:\test.xlsx", "test 4", "B", 1);

            OfflineExcel.DeleteTextFromCell(@"E:\test.xlsx", "test", "B", 1);
        }
    }
}

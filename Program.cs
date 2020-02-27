using System;
using System.IO;

namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = Path.Combine(Directory.GetCurrentDirectory(), "TestExcel.xlsx");
            using (var excel = new Excel(fileName))
            {
                var inputString = "TestValue";
                excel.WriteToCell(1, 1, inputString);
                var value = excel.ReadCell(1, 1);
                Console.WriteLine(value);
                excel.Save();
                excel.Close();
            }
        }
    }
}

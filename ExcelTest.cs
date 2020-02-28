using System.IO;
using NUnit.Framework;

namespace Excel
{
    [TestFixture]
    public class ExcelTest
    {
        [Test]
        public void ReadWriteTest()
        {
            var inputString = "TestValue";
            string readString;
            var fileName = Path.Combine(TestContext.CurrentContext.TestDirectory, "TestExcel.xlsx");
            using (var excel = new Excel(fileName))
            {
                excel.WriteToCell(1, 1, inputString);
                readString = excel.ReadCell(1, 1);
                excel.Save();
                excel.Close();
            }

            Assert.AreEqual(inputString, readString);
        }
    }
}

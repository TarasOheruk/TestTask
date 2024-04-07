using Microsoft.VisualStudio.TestPlatform.TestHost;
using NUnit.Framework.Internal.Execution;
using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;

namespace TestProject1
{
    [TestFixture]
    public class UnitTest1
    {
        private Program _program { get; set; } = null!;

        private const string TestExcelFileName = "TestExcel.xlsx";
        private const string TestCSVFileName = "TestOutput.csv";

        [SetUp]
        public void SetUp()
        {
            _program = new Program();
        }
        [Test]
        public async Task TestGetExcelFileUrl()
        {
            // Arrange
            string html = "<html><body><a title='Worldwide Rig Count Jan 2007_Mar 2024.xlsx' href='https://example.com/excel.xlsx'>Link</a></body></html>";

            // Act
            string url = await Program.GetExcelFileUrl(html);

            // Assert
            Assert.That(url, Is.EqualTo("https://example.com/excel.xlsx"));
        }
        [Test]
        public async Task TestConvertExcelToCSV_FileExistsAndHasWorksheet_ConvertsSuccessfully()
        {
            string fileName = "test.xlsx";
            string outputFileName = "test.csv";

            var excelPackage = new ExcelPackage();
            var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
            worksheet.Cells["A1"].Value = "Data1";
            worksheet.Cells["B1"].Value = "Data2";
            worksheet.Cells["A2"].Value = "Data3";
            worksheet.Cells["B2"].Value = "Data4";

            await excelPackage.SaveAsAsync(new FileInfo(fileName));

            await Program.ConvertExcelToCSV(fileName, outputFileName);

            Assert.That(File.Exists(outputFileName), Is.True);

        }

        [Test]
        public async Task TestGetRowData()
        {
            using (var package = new OfficeOpenXml.ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells[1, 1].Value = "Test";

                var rowData = await Program.GetRowData(worksheet, 1, 1);

                Assert.That(rowData, Is.EqualTo("Test"));
            }
        }
    }
}
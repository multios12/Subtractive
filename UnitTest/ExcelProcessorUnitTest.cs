namespace UnitTest
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Subtractive.Processor;
    using System.IO;

    [TestClass]
    public class ExcelProcessorUnitTest
    {
        ExcelProcessor processor = new ExcelProcessor();

        [DataTestMethod]
        [DataRow(@"Resources\ExcelPng.xlsm")]
        [DataRow(@"Resources\ExcelPng.xlsx")]
        [DataRow(@"Resources\ExcelJpeg.xlsm")]
        [DataRow(@"Resources\ExcelJpeg.xlsx")]
        public void ExecuteTest(string filePath)
        {
            processor.Execute(filePath);
            string distPath = Path.Combine("Resources", "減色済" + Path.GetFileName(filePath));
            Assert.IsTrue(File.Exists(distPath));
        }
    }
}

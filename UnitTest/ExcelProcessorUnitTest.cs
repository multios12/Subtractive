namespace UnitTest
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Subtractive.Processor;
    using System.IO;

    [TestClass]
    public class ExcelProcessorUnitTest
    {
        [TestMethod,TestCategory("ExcelProcessor")]
        public void ExecuteTestFromPngXlsx()
        {
            ExcelProcessor processor = new ExcelProcessor();
            processor.Execute(@"Resources\pngBook.xlsx");
            Assert.IsTrue(File.Exists(@"Resources\減色済pngBook.xlsx"));
        }

        [TestMethod, TestCategory("ExcelProcessor")]
        public void ExecuteTestFromPngXlsm()
        {
            ExcelProcessor processor = new ExcelProcessor();
            processor.Execute(@"Resources\pngBook.xlsm");
            Assert.IsTrue(File.Exists(@"Resources\減色済pngBook.xlsm"));
        }
    }
}

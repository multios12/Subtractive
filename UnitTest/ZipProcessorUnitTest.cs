namespace UnitTest
{
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Subtractive.Processor;

    [TestClass]
    public class ZipProcessorUnitTest
    {
        ZipProcessor processor = new ZipProcessor() { IsConvertToPng = true };

        [DataTestMethod]
        [DataRow(@"Resources\jpeg01.zip")]
        [DataRow(@"Resources\png01.zip")]
        public void ExecuteTest(string filePath)
        {
            Assert.IsNotNull(processor.ResizeSize);
            Assert.AreEqual<int>(800, processor.ResizeSize.Value.Height);
            Assert.AreEqual<int>(1400, processor.ResizeSize.Value.Width);

            processor.Execute(filePath);

            string distPath = Path.Combine("Resources", "減色済" + Path.GetFileName(filePath));
            Assert.IsTrue(File.Exists(distPath));
        }
    }
}

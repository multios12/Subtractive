namespace UnitTest
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Subtractive.Processor;
    using System.IO;

    [TestClass]
    public class WordProcessorUnitTest
    {
        WordProcessor processor = new WordProcessor();



        [DataTestMethod]
        [DataRow(@"Resources\WordPng.docm")]
        [DataRow(@"Resources\WordPng.docx")]
        [DataRow(@"Resources\WordJpeg.docm")]
        [DataRow(@"Resources\WordJpeg.docx")]
        public void ExecuteTest(string filePath)
        {
            processor.Execute(filePath);
            string distPath = Path.Combine("Resources", "減色済" + Path.GetFileName(filePath));
            Assert.IsTrue(File.Exists(distPath));
        }
    }
}

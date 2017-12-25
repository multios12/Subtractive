namespace UnitTest
{
    using System;
    using System.IO;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Subtractive.Processor;

    [TestClass]
    public class ZipProcessorUnitTest
    {
        ZipProcessor processor = new ZipProcessor();

        [TestMethod]
        public void ExecuteTestFromPng()
        {
            processor.Execute(@"Resources\jpeg01.zip");
            Assert.IsTrue(File.Exists(@"Resources\減色済jpeg01.zip"));
        }

        [TestMethod]
        public void ExecuteTestFromJpeg()
        {
            processor.Execute(@"Resources\png01.zip");
            Assert.IsTrue(File.Exists(@"Resources\減色済png01.zip"));
        }
    }
}

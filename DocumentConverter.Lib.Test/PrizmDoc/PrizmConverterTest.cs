using Accusoft.PrizmDoc;
using Accusoft.PrizmDocServer.Conversion;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace DocumentConverter.Lib.Test.PrizmDoc
{
    /// <summary>
    /// Summary description for PrizmConverterTest
    /// </summary>
    [TestClass]
    public class PrizmConverterTest
    {
        private const string apiKey = "MjXoHvf0ia9tjXOpFUgCny2XWEYOeQIOiLb3MIY0Nqu4XA7qzGxLm7jqSS8ykQPa";
        private readonly IDocConverter _docConverter;
        public PrizmConverterTest()
        {
            _docConverter = new PrizmDocConverter(apiKey);
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public async Task ConvertRtfToDox_Test()
        {
            const string filename = "largest_chapter.html";
            string sourceFilePath = FileManager.GetPath(filename);
            string destinationFilePath = 
                FileManager.FilePathToWriteTo($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_{filename.Replace("html", "pdf")}","prizmdoc");
            await _docConverter.ConvertRtfToDocxAsync(sourceFilePath, destinationFilePath);
            Assert.IsTrue(File.Exists(destinationFilePath));
        }
    }
}

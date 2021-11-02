using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading;

namespace DocumentConverter.Lib.Test
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass,TestCategory("Cloudmersive")]
    public class DocumentConverterTest
    {
        IDocumentConverter _documentConverter;
        const string cloudMersiveApiKey = "b526286d-3567-4369-af59-dfccb0801b9b";
        public DocumentConverterTest()
        {
            _documentConverter = new DocumentConverter(cloudMersiveApiKey);
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
        [TestInitialize()]
        public void MyTestInitialize() {
           
        }
        //
        //Use TestCleanup to run code after each test has run
        [TestCleanup()]
         public void MyTestCleanup() {
        }
        //
        #endregion
        [TestMethod]
        public void ConvertDoxToPdf_Test()
        {
            string filename = "324376.docx";
            var stream = FileManager.LoadFile(filename);

            var watch = new Stopwatch();
            watch.Start();
            var data = _documentConverter.ConvertDocToPdf(stream);
            watch.Stop();
            Assert.IsNotNull(data);

            Debug.WriteLine($"Execution Time (ConvertDocToPdf): {watch.Elapsed.TotalSeconds} s");

            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_{filename.Replace("docx", "pdf")}", data)));
        }
        [TestMethod]
        public void ConvertPdfToDocx_Test()
        {
            string filename = "324376.pdf";
            var stream = FileManager.LoadFile(filename);

            var watch = new Stopwatch();
            watch.Start();
            var data = _documentConverter.ConvertPdfToDocx(stream);
            watch.Stop();
            Assert.IsNotNull(data);

            Debug.WriteLine($"Execution Time (ConvertPdfToDocx): {watch.Elapsed.TotalSeconds} s");

            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_{filename.Replace("pdf", "docx")}", data)));
        }
        [TestMethod]
        public void ConvertRtfToDocx_Test()
        {
            string filename = "41287.rtf";
            var stream = FileManager.LoadFile(filename);

            var watch = new Stopwatch();
            watch.Start();
            var data = _documentConverter.ConvertRtfToDocx(stream);
            watch.Stop();
            Assert.IsNotNull(data);
           
            Debug.WriteLine($"Execution Time (ConvertRtfToDocx): {watch.Elapsed.TotalSeconds} s");
           
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_{filename.Replace("rtf", "docx")}", data)));
        }
        [TestMethod]
        public void ConvertHtmlToPdf_Test()
        {
            string filename = "largest_chapter.html";
            var stream = FileManager.LoadFile(filename);

            var watch = new Stopwatch();
            watch.Start();
            var data = _documentConverter.ConvertHtmlToPdf(stream);
            watch.Stop();
            Assert.IsNotNull(data);

            Debug.WriteLine($"Execution Time (ConvertHtmlToPdf): {watch.Elapsed.TotalSeconds} s");

            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_{filename.Replace("html", "pdf")}", data)));
        }
        [TestMethod]
        public void ConvertDocumentPptxToPdf_Test()
        {
            string filename = "08-10-2021_0910_129458_Poll_Charles_R1.pptx";
            var inputFile = FileManager.LoadFile(filename);

            var watch = new Stopwatch();
            watch.Start();
            var data = _documentConverter.ConvertDocumentPptxToPdf(inputFile);
            watch.Stop();
            Assert.IsNotNull(data);

            Debug.WriteLine($"Execution Time (ConvertDocumentPptxToPdf): {watch.Elapsed.TotalSeconds} s");

            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_{filename.Replace("pptx", "pdf")}", data)));
        }
        [TestMethod]
        public void ConvertDocumentPptxToPng_Test()
        {
            string filename = "08-10-2021_0910_129458_Poll_Charles_R2.pptx";
            var inputFile = FileManager.LoadFile(filename);
            var watch = new Stopwatch();
            watch.Start();
            var data = _documentConverter.ConvertDocumentPptxToPng(inputFile);
            watch.Stop();
            Assert.IsNotNull(data);

            Debug.WriteLine($"Execution Time (ConvertDocumentPptxToPng): {watch.Elapsed.TotalSeconds} s");
            foreach (var image in data.PngResultPages)
            {

                FileManager.DownloadFile(image.URL);
            }
           // Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_{filename.Replace("pptx", "pdf")}", data)));
        }
    }
}

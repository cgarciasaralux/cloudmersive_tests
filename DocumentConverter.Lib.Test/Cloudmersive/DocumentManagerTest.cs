using DocumentConverter.Lib.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
namespace DocumentConverter.Lib.Test
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass, TestCategory("Cloudmersive")]
    public class DocumentManagerTest
    {
        IDocumentManager _documentManager;
        const string cloudMersiveApiKey = "b526286d-3567-4369-af59-dfccb0801b9b";
        public DocumentManagerTest()
        {
            _documentManager = new DocumentManager(cloudMersiveApiKey);
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


        //Some performance improvement because creating HTML from string
        [TestMethod]
        [TestCategory("HTML->DOCX")]
        public async Task CreateFormCoverPage_Test()
        {
            FormCitation formCitation = new FormCitation {
             PlusLogoUrl = "https://plus.pli.edu/Content/css/images/PLI_PLUS_LOGO.png",
             FormTitle= "Sample Jury Instructions from Reported Decisions",
             BookTitle= "53rd Annual Immigration and Naturalization Institute",
             PermaLink= "https://plus.pli.edu/Details/Details?fq=id:(41509)",
            };

            var watch = new Stopwatch();
            watch.Start();
            var data = await _documentManager.CreateFormCoverPage(formCitation);
            //var data = await _documentManager.CreateFormCoverPageAsync(formCitation);
            watch.Stop();
            Assert.IsNotNull(data);
            Debug.WriteLine($"Execution Time (CreateFormCoverPage): {watch.Elapsed.TotalSeconds} s");
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_FormCoverPage.docx", data)));

        }

        [TestMethod]
        [TestCategory("HTML->DOCX")]
        public async Task CreateFormCoverPage_Test_Async()
        {
            //ARRANGE
            string htmlTemplate = FileManager.ReadFileAsString("Form Template.html");

            FormCitation formCitation = new FormCitation
            {
                PlusLogoUrl = "https://plus.pli.edu/Content/css/images/PLI_PLUS_LOGO.png",
                FormTitle = "Sample Jury Instructions from Reported Decisions",
                BookTitle = "53rd Annual Immigration and Naturalization Institute",
                PermaLink = "https://plus.pli.edu/Details/Details?fq=id:(41509)",
            };

            //ACT
            var watch = new Stopwatch();
            watch.Start();
            var data = await _documentManager.CreateFormCoverPageAsync(formCitation, htmlTemplate);
            watch.Stop();

            //ASSERT
            Assert.IsNotNull(data);
            Debug.WriteLine($"Execution Time (CreateFormCoverPageAsync): {watch.Elapsed.TotalSeconds} s");
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_FormCoverPageAsync.docx", data)));

        }

        //Some improvement, but not much, by using async programming
        [TestMethod]
        [TestCategory("Word template")]
        public void CreateFormCoverPageTemplate_Test()
        {
            //ARRANGE
            string templateFileName = "FormTemplate.docx";
            FormCitation formCitation = new FormCitation
            {
                PlusLogoUrl = "https://plus.pli.edu/Content/css/images/PLI_PLUS_LOGO.png",
                FormTitle = "Sample Jury Instructions from Reported Decisions",
                BookTitle = "53rd Annual Immigration and Naturalization Institute",
                PermaLink = "https://plus.pli.edu/Details/Details?fq=id:(41509)",
            };
            var file = FileManager.ReadFile(templateFileName);

            //ACT
            var watch = new Stopwatch();
            watch.Start();
            var data = _documentManager.CreateFormCoverPageTemplate(formCitation,file);
            //var data = _documentManager.CreateFormCoverPageTemplateAsync(formCitation, file);
            watch.Stop();

            //ASSERT
            Assert.IsNotNull(data);
            Debug.WriteLine($"Execution Time (CreateFormCoverPageTemplate): {watch.Elapsed.TotalSeconds} s");
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_FormCoverPageTemplate.docx", data.Result)));

        }

        [TestMethod]
        [TestCategory("Word template")]
        public void CreateFormCoverPageTemplate_Test_Async()
        {
            //ARRANGE
            string templateFileName = "FormTemplate.docx";
            FormCitation formCitation = new FormCitation
            {
                PlusLogoUrl = "https://plus.pli.edu/Content/css/images/PLI_PLUS_LOGO.png",
                FormTitle = "Sample Jury Instructions from Reported Decisions",
                BookTitle = "53rd Annual Immigration and Naturalization Institute",
                PermaLink = "https://plus.pli.edu/Details/Details?fq=id:(41509)",
            };
            var file = FileManager.ReadFile(templateFileName);

            //ACT
            var watch = new Stopwatch();
            watch.Start();
            //var data = _documentManager.CreateFormCoverPageTemplate(formCitation,file);
            var data = _documentManager.CreateFormCoverPageTemplateAsync(formCitation, file);
            watch.Stop();

            //ASSERT
            Assert.IsNotNull(data);
            Debug.WriteLine($"Execution Time (CreateFormCoverPageTemplateAsync): {watch.Elapsed.TotalSeconds} s");
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_FormCoverPageTemplateAsync.docx", data.Result)));

        }


        [TestMethod]
        public void MergeDocx_Test()
        {
            string[] files = new string[] { "324380.docx", "324376.docx", "324381.docx" };
            FormCitation formCitation = new FormCitation
            {
                PlusLogoUrl = "https://plus.pli.edu/Content/css/images/PLI_PLUS_LOGO.png",
                FormTitle = "Sample Jury Instructions from Reported Decisions",
                BookTitle = "53rd Annual Immigration and Naturalization Institute",
                PermaLink = "https://plus.pli.edu/Details/Details?fq=id:(41509)",
            };
            foreach (var f in files)
            {
                var watch = new Stopwatch();
                watch.Start();
                var formCoverPage = _documentManager.CreateBlankDoc();//.CreateFormCoverPage(formCitation);
                var form = FileManager.LoadFile(f);
                var name = f.Substring(0, f.IndexOf("."));
             
                
                var data = _documentManager.MergeDocx(new MemoryStream(formCoverPage), form);
                watch.Stop();
                Assert.IsNotNull(data);
                Debug.WriteLine($"Execution Time (MergeDocx) Test {name}: {watch.Elapsed.TotalSeconds} s");
                Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_{name}_MergeDocx.docx", data)));
            }
           
        }

        [TestMethod]
        public void CreateBlankDocx_Test()
        {
            var watch = new Stopwatch();
            watch.Start();
            var blankDoc = _documentManager.CreateBlankDoc();
            watch.Stop();
            Debug.WriteLine($"Execution Time (CreateBlankDoc) Test: {watch.Elapsed.TotalSeconds} s");
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_CreateBlankDoc.docx", blankDoc)));

        }
        
        //Some improvement when using HTML from string
        [TestMethod]
        [TestCategory("HTML")]
        public async Task CreateFormCoverPageHtml_Test()
        {
            //ARRANGE
            FormCitation formCitation = new FormCitation
            {
                PlusLogoUrl = "https://plus.pli.edu/Content/css/images/PLI_PLUS_LOGO.png",
                FormTitle = "Sample Jury Instructions from Reported Decisions",
                BookTitle = "53rd Annual Immigration and Naturalization Institute",
                PermaLink = "https://plus.pli.edu/Details/Details?fq=id:(41509)",
            };

            //ACT
            var watch = new Stopwatch();
            watch.Start();
            var data = await _documentManager.CreateFormCoverPageHtml(formCitation);
            //var data = await _documentManager.CreateFormCoverPageHtmlAsync(formCitation);
            watch.Stop();

            //ASSERT
            Assert.IsNotNull(data);
            Debug.WriteLine($"Execution Time (CreateFormCoverPageHtml): {watch.Elapsed.TotalSeconds} s");
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_FormCoverPageHtml.html", data)));

        }

        [TestMethod]
        [TestCategory("HTML")]
        public async Task CreateFormCoverPageHtml_Test_Async()
        {
            //ARRANGE
            string htmlTemplate = FileManager.ReadFileAsString("Form Template.html");

            FormCitation formCitation = new FormCitation
            {
                PlusLogoUrl = "https://plus.pli.edu/Content/css/images/PLI_PLUS_LOGO.png",
                FormTitle = "Sample Jury Instructions from Reported Decisions",
                BookTitle = "53rd Annual Immigration and Naturalization Institute",
                PermaLink = "https://plus.pli.edu/Details/Details?fq=id:(41509)",
            };

            //ACT
            var watch = new Stopwatch();
            watch.Start();
            //var data = await _documentManager.CreateFormCoverPageHtml(formCitation);
            var data = await _documentManager.CreateFormCoverPageHtmlAsync(formCitation, htmlTemplate);
            watch.Stop();

            //ASSERT
            Assert.IsNotNull(data);
            Debug.WriteLine($"Execution Time (CreateFormCoverPageHtmlAsync): {watch.Elapsed.TotalSeconds} s");
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_FormCoverPageHtmlAsync.html", data)));

        }

        //try to use HTML instead DOCX but its the same or slower
        [TestMethod]
        public void CreatePdfCoverPage_Test()
        {
            //ARRANGE
            Stream inputFile2 = FileManager.LoadFile("15140-CH11.pdf");

            //ACT
            var watch = new Stopwatch();
            watch.Start();
            var inputFile1 = _documentManager.CreatePDFCoverPage();
            //var inputFile1 = _documentManager.CreatePDFCoverPageAsync();
            var data = _documentManager.MergePdfs(new MemoryStream(inputFile1), inputFile2);
            watch.Stop();

            //ASSERT
            Assert.IsNotNull(data);
            Debug.WriteLine($"Execution Time (CreatePDFCoverPage): {watch.Elapsed.TotalSeconds} s");
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_Pdf_With_CoverPage.pdf", data)));
        }


        [TestMethod]
        public async Task AddPdfAnnotations_Test()
        {
            //ARRANGE
            byte[] inputData = FileManager.ReadFile("15140-CH11.pdf");
            byte[] annotationsInput = FileManager.ReadFile("Annotations.json");
            var annotations = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Annotation>>(
               Encoding.UTF8.GetString(annotationsInput));
            
            //ACT
            var watch = new Stopwatch();
            watch.Start();
            var data = await _documentManager.AddPdfAnnotations(inputData, annotations);
            watch.Stop();

            //ASSERT
            Assert.IsNotNull(data);
            Debug.WriteLine($"Execution Time (GetPdfAnnotations): {watch.Elapsed.TotalSeconds} s");
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_Pdf_With_Added_Annotations.pdf", data)));
        }

        [TestMethod]
        public void GetPdfAnnotations_Test()
        {
            Stream inputFile = FileManager.LoadFile("15140-CH11_with_annotations.pdf");
            var watch = new Stopwatch();
            watch.Start();
            var data = _documentManager.GetPdfAnnotations(inputFile);
            watch.Stop();
            Assert.IsTrue(data.Count > 0);
            Debug.WriteLine($"Execution Time (GetPdfAnnotations): {watch.Elapsed.TotalSeconds} s");
            string json = Newtonsoft.Json.JsonConvert
                            .SerializeObject(data);
            Assert.IsTrue(File.Exists(FileManager.WriteToFile($"{DateTimeOffset.Now.ToUnixTimeSeconds()}_Annotations.json",Encoding.UTF8.GetBytes
                (json))));
        }
    }
}

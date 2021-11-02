using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Api;
using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Client;
using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Model;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
namespace Cloudmersive.Test.Console
{
    class Program
    {
        static string blob =
            "This form was submitted as part of the Course Handbook for <b>53rd Annual Immigration and Naturalization" +
            " Institute</b> (Copyright © 2021 by Practising Law Institute. All rights reserved.)" +
            " Use of the form is subject to the copyright of the author(s) and/or licensee(s)," +
            " and is made available here only for use in the normal course of business," +
            " including for legal research, drafting, business development " +
            "and the provision of other legal services. Questions can be directed to " +
            "<a href=\"mailto:PLUS@pli.edu\">PLUS@pli.edu</a>";
        static void Main(string[] args)
        {
            var listeners = new TraceListener[] { new TextWriterTraceListener(System.Console.Out) };
            Debug.Listeners.AddRange(listeners);
            //To get the location the assembly normally resides on disk or the install directory
            string path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            // Configure API key authorization: Apikey
            Configuration.Default.AddApiKey("Apikey", "7908739d-d9df-448d-afee-a9ec0445b715");


            // RenderForm();
            ConvertRtfToWord("41287.rtf");

        }

        private static void ConvertRtfToWord(string fileName)
        {
            string path = $"{ Directory.GetCurrentDirectory() }/Data/{fileName}";
            FileInfo fi = new FileInfo(fileName);
            ConvertDocumentApi convertDocumentApi = new ConvertDocumentApi();

            using (FileStream stream = new FileStream(path, FileMode.Open))
            {

                Stopwatch watch = new Stopwatch();
                watch.Start();
                var bytes = convertDocumentApi.ConvertDocumentRtfToDocx(stream);
                watch.Stop();
                Debug.WriteLine($"Execution Time (ConvertRtfToWord): {watch.ElapsedMilliseconds} ms FileSize:{fileName.Length} bytes");
                WriteToFile($"{fileName}_{DateTimeOffset.Now.ToUnixTimeMilliseconds()}.docx", bytes);
            }
          

        }

        private static void RenderForm()
        {
            byte[] result = CreateFormCoverPage();
            MergeDocumentApi mergeDocumentApi = new MergeDocumentApi();
            byte[] formResult = GetDocument("324380.docx");

            try
            {
                var watch = new Stopwatch();
                watch.Start();
                var mergeResult = mergeDocumentApi.MergeDocumentDocx(new MemoryStream(result), new MemoryStream(formResult));
                watch.Stop();
                Debug.WriteLine($"Execution Time (MergeDocumentDocx): {watch.ElapsedMilliseconds} ms");

                WriteToFile($"324380_{DateTimeOffset.Now.ToUnixTimeMilliseconds()}.docx",mergeResult);

            }
            catch (Exception ex)
            {

                Debug.WriteLine(ex.Message);
            }
        }

        private static void WriteToFile(string filename, byte[] mergeResult)
        {
            using (Stream file = File.OpenWrite($@"C:\Users\bhassan\Downloads\cloudmersive\{filename}"))
            {
                file.Write(mergeResult, 0, mergeResult.Length);
            }
        }

        private static byte[] GetDocument(string name)
        {
            string path =  $"{ Directory.GetCurrentDirectory() }/Data/{name}";
            byte[] bytes = System.IO.File.ReadAllBytes(path);
            return bytes;
        }

        private static byte[] CreateFormCoverPage()
        {
            var watch = new Stopwatch();
            watch.Start();

            var editHtmlApi = new EditHtmlApi();
            var convertHtml = new ConvertWebApi();
            var htmlDocument = editHtmlApi.EditHtmlHtmlCreateBlankDocument("Form Cover Page");
            htmlDocument = editHtmlApi.EditHtmlHtmlAppendImageFromUrl(
                "https://plus.pli.edu/Content/css/images/PLI_PLUS_LOGO.png", new MemoryStream(htmlDocument));
            HtmlToOfficeRequest htmlToOfficeRequest = new HtmlToOfficeRequest();

            var hr_paragraph = "<hr/>";

            htmlDocument = editHtmlApi.EditHtmlHtmlAppendParagraph(hr_paragraph, new MemoryStream(htmlDocument));

            /*
             Form Name: Sample Jury Instructions from Reported Decisions
               URL: https://plus.pli.edu/Details/Details?fq=id:(41509)

             */

            var formName = "<b>Form Name: </b><span>Sample Jury Instructions from Reported Decisions</span><br/>";
            var formLink = "<b>Url: </b><a href=\"https://plus.pli.edu/Details/Details?fq=id:(41509)\">https://plus.pli.edu/Details/Details?fq=id:(41509)</a><br/>";


            htmlDocument = editHtmlApi.EditHtmlHtmlAppendParagraph(formName + formLink, new MemoryStream(htmlDocument));

            htmlDocument = editHtmlApi.EditHtmlHtmlAppendParagraph(blob, new MemoryStream(htmlDocument));

            htmlToOfficeRequest.Html = Encoding.UTF8.GetString(htmlDocument, 0, htmlDocument.Length);

            var result =  convertHtml.ConvertWebHtmlToDocx(htmlToOfficeRequest);
            watch.Stop();

            Debug.WriteLine($"Execution Time (CreateFormCoverPage): {watch.ElapsedMilliseconds} ms");
            return result;
        }
        #region old
        private static byte[] CoverPage(FileStream formTemplateInput)
        {
            var editDocumentApi = new EditDocumentApi();
            
            var documentUrl = editDocumentApi.EditDocumentBeginEditing(formTemplateInput).Replace("\"", "");


            ReplaceStringRequest replaceTitle = new ReplaceStringRequest(
                inputFileUrl: documentUrl,
                replaceString: "Current Program Electronic Review Management Audit Language",
                matchCase: false,
                matchString: "[name]"
                );

           var result1 = editDocumentApi.EditDocumentDocxReplace(replaceTitle);

            ReplaceStringRequest replaceUrl = new ReplaceStringRequest(
              inputFileBytes: result1,
              replaceString: "https://plus.pli.edu/Details/Details?fq=id:(324381)",
              matchCase: false,
              matchString: "[Url]"
              );

            var result2 = editDocumentApi.EditDocumentDocxReplace(replaceUrl);


            ReplaceStringRequest replaceEmail = new ReplaceStringRequest(
              inputFileBytes: result2,
             replaceString: "Plus@pli.edu",
             matchCase: false,
             matchString: "[Email]"
             );

            var result3 = editDocumentApi.EditDocumentDocxReplace(replaceEmail);


            ReplaceStringRequest replaceYear = new ReplaceStringRequest(
             inputFileBytes: result3,
            replaceString: DateTime.Now.Year.ToString(),
            matchCase: false,
            matchString: "[Year]"
            );

            return editDocumentApi.EditDocumentDocxReplace(replaceYear);

            // DocxParagraph formNameParagraph = new DocxParagraph();
            // formNameParagraph.ContentRuns = new List<DocxRun> {
            //      CreateLabel("Form Name:"),
            //      CreateContent("Current Program Electronic Review Management Audit Language") };
            // formNameParagraphRequest.ParagraphToInsert = formNameParagraph;



            // InsertDocxInsertParagraphResponse formNameParagraphResponse =
            //     editDocumentApi.EditDocumentDocxInsertParagraph(formNameParagraphRequest);


            // InsertDocxInsertParagraphRequest formUrlParagraphRequest = new InsertDocxInsertParagraphRequest();
            // formUrlParagraphRequest.InputFileUrl = formNameParagraphResponse.EditedDocumentURL;
            // DocxParagraph formUrlParagraph = new DocxParagraph();
            // formUrlParagraph.ContentRuns = new List<DocxRun> {
            //      CreateLabel("Url:"),
            //      CreateContent( "https://plus.pli.edu/Details/Details?fq=id:(324381)\n") };
            // formUrlParagraphRequest.ParagraphToInsert = formUrlParagraph;

            // InsertDocxInsertParagraphResponse formUrlParagraphResponse =
            //     editDocumentApi.EditDocumentDocxInsertParagraph(formUrlParagraphRequest);

            // InsertDocxInsertParagraphRequest blobParagraphRequest = new InsertDocxInsertParagraphRequest();
            // blobParagraphRequest.InputFileUrl = formUrlParagraphResponse.EditedDocumentURL;
            // DocxParagraph blobParagraph = new DocxParagraph();
            // blobParagraph.ContentRuns = new List<DocxRun> {
            //      CreateLabel(""),
            //      CreateContent(blob) };
            // blobParagraphRequest.ParagraphToInsert = blobParagraph;

            // InsertDocxInsertParagraphResponse blobParagraphResponse =
            //editDocumentApi.EditDocumentDocxInsertParagraph(blobParagraphRequest);

            FinishEditingRequest finishEditingRequest = new FinishEditingRequest(documentUrl);
            var result = editDocumentApi.EditDocumentFinishEditing(finishEditingRequest);
            return result;
        }
        private static DocxRun CreateContent(string content)
        {
            //title
            DocxRun nameRun = new DocxRun();
            DocxText bookTitle = new DocxText { TextContent = content };
            nameRun.TextItems = new List<DocxText> { bookTitle };

            return nameRun;
        }

        private static DocxRun CreateLabel(string lblText)
        {
            //Form label
            DocxRun lblNameRun = new DocxRun();

            DocxText lblName = new DocxText { TextContent = lblText };
            lblNameRun.Bold = true;
            lblNameRun.TextItems = new List<DocxText> { lblName };
            return lblNameRun;
        }
        #endregion
    }
}

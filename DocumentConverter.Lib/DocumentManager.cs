using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Api;
using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Client;
using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Model;
using DocumentConverter.Lib.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentConverter.Lib
{
    public class DocumentManager : IDocumentManager
    {

        public DocumentManager(string cloudMersiveApiKey)
        {
            Configuration.Default.AddApiKey("Apikey",cloudMersiveApiKey);
        }

        public async Task<byte[]> CreateFormCoverPageHtmlAsync(FormCitation formCitation, string template)
        {
            string html = GetHtmlTemplateAsString(template, "Form Cover Page", formCitation.PlusLogoUrl, formCitation.FormTitle, formCitation.PermaLink, "PLUS@pli.edu");

            Task<byte[]> task = Task<byte[]>.Factory.StartNew(() =>
            {
                return Encoding.ASCII.GetBytes(html);
            });

            return await task;
        }

        private string GetHtmlTemplateAsString(string html, string title, string logoUrl, string formTitle, string permaLink, string email)
        {
            return string.Format(html, title, logoUrl, formTitle, permaLink, DateTime.Now.Year, email);
        }

        public async Task<byte[]> CreateFormCoverPageHtml(FormCitation formCitation)
        {
            EditHtmlApi editHtmlApi = new EditHtmlApi();
            Task<byte[]> htmlDocument = editHtmlApi.EditHtmlHtmlCreateBlankDocumentAsync("Form Cover Page");

            //htmlDocument = editHtmlApi.EditHtmlHtmlAppendImageFromUrlAsync(formCitation.PlusLogoUrl, new MemoryStream(htmlDocument.Result));
            //htmlDocument = editHtmlApi.EditHtmlHtmlAppendImageFromUrlAsync(formCitation.PlusLogoUrl);
            /*
              Form Name: Sample Jury Instructions from Reported Decisions
              URL: https://plus.pli.edu/Details/Details?fq=id:(41509)
            */
            StringBuilder sb = new StringBuilder();
            sb.Append($"<hr/><p><b>Form Name: </b><span>{formCitation.FormTitle}</span><br/>");
            sb.Append($"<b>Url: </b><a href=\"{formCitation.PermaLink}\">{formCitation.PermaLink}</a><br/></p>");
            sb.Append($"<p>This form was submitted as part of the Course Handbook for <b>{formCitation.BookTitle}</b>" +
                $" (Copyright © {DateTime.Now.Year} by Practising Law Institute. All rights reserved.)" +
            " Use of the form is subject to the copyright of the author(s) and/or licensee(s)," +
            " and is made available here only for use in the normal course of business," +
            " including for legal research, drafting, business development " +
            "and the provision of other legal services. Questions can be directed to " +
            "<a href=\"mailto:PLUS@pli.edu\">PLUS@pli.edu</a></p>");

            return await editHtmlApi.EditHtmlHtmlAppendParagraphAsync(sb.ToString(), new MemoryStream(htmlDocument.Result));
        }
        
        public async Task<byte[]> CreateFormCoverPageAsync(FormCitation formCitation, string template)
        {
            string html = GetHtmlTemplateAsString(template, "Form Cover Page", formCitation.PlusLogoUrl, formCitation.FormTitle, formCitation.PermaLink, "PLUS@pli.edu");

            return await new ConvertWebApi()
                .ConvertWebHtmlToDocxAsync(new HtmlToOfficeRequest(html));
        }

    

        public async Task<byte[]> CreateFormCoverPage(FormCitation formCitation)
        {
            ConvertWebApi convertHtml = new ConvertWebApi();
            HtmlToOfficeRequest htmlToOfficeRequest = new HtmlToOfficeRequest();
            var htmlDocument = await CreateFormCoverPageHtml(formCitation);

            htmlToOfficeRequest.Html = Encoding.UTF8.GetString(htmlDocument, 0, htmlDocument.Length);
            var result = await convertHtml.ConvertWebHtmlToDocxAsync(htmlToOfficeRequest);
            return result;
        }

        public async Task<byte[]> CreateFormCoverPageTemplateAsync(FormCitation formCitation, byte[] formTemplateInput)
        {
            Task<SingleReplaceString> replaceTitleTask = GetReplaceStringAsync(formCitation.FormTitle, "[name]");
            Task<SingleReplaceString> replaceUrlTask = GetReplaceStringAsync(formCitation.PermaLink, "[Url]");
            Task<SingleReplaceString> replaceEmailTask = GetReplaceStringAsync("Plus@pli.edu", "[Email]");
            Task<SingleReplaceString> replaceYearTask = GetReplaceStringAsync(DateTime.Now.ToString(), "[Year]");

            SingleReplaceString[] results = await Task.WhenAll(replaceTitleTask, replaceUrlTask, replaceEmailTask, replaceYearTask);
            
            MultiReplaceStringRequest multiReplaceStringRequest = new MultiReplaceStringRequest { };
            multiReplaceStringRequest.ReplaceStrings = results.ToList();
            multiReplaceStringRequest.InputFileBytes = formTemplateInput;

            EditDocumentApi editDocumentApi = new EditDocumentApi();
            return await editDocumentApi.EditDocumentDocxReplaceMultiAsync(multiReplaceStringRequest);
        }

        private async Task<SingleReplaceString> GetReplaceStringAsync(string replaceString, string matchString)
        {
            Task<SingleReplaceString> task = Task<SingleReplaceString>.Factory.StartNew(() =>
            {
                return new SingleReplaceString(
                    replaceString: replaceString,
                    matchCase: false,
                    matchString: matchString
                );
            });

            return await task;
        }

        public async Task<byte[]> CreateFormCoverPageTemplate(FormCitation formCitation, byte[] formTemplateInput)
        {
            var editDocumentApi = new EditDocumentApi();
            //var documentUrl = editDocumentApi.EditDocumentBeginEditing(formTemplateInput).Replace("\"", "");

            MultiReplaceStringRequest multiReplaceStringRequest = new MultiReplaceStringRequest { };

            SingleReplaceString replaceTitle = new SingleReplaceString(
                //inputFileUrl: documentUrl,
                replaceString: formCitation.FormTitle,
                matchCase: false,
                matchString: "[name]"
                );

            multiReplaceStringRequest.ReplaceStrings = new List<SingleReplaceString>();
            multiReplaceStringRequest.ReplaceStrings.Add(replaceTitle);
            // var result1 = editDocumentApi.EditDocumentDocxReplace(replaceTitle);

            SingleReplaceString replaceUrl = new SingleReplaceString(
             // inputFileBytes: result1,
              replaceString: formCitation.PermaLink,
              matchCase: false,
              matchString: "[Url]"
              );

            //var result2 = editDocumentApi.EditDocumentDocxReplace(replaceUrl);
            multiReplaceStringRequest.ReplaceStrings.Add(replaceUrl);

            SingleReplaceString replaceEmail = new SingleReplaceString(
             // inputFileBytes: result2,
             replaceString: "Plus@pli.edu",
             matchCase: false,
             matchString: "[Email]"
             );

            //  var result3 = editDocumentApi.EditDocumentDocxReplace(replaceEmail);
            multiReplaceStringRequest.ReplaceStrings.Add(replaceEmail);

            SingleReplaceString replaceYear = new SingleReplaceString(
           //  inputFileBytes: result3,
            replaceString: DateTime.Now.Year.ToString(),
            matchCase: false,
            matchString: "[Year]"
            );
            multiReplaceStringRequest.ReplaceStrings.Add(replaceYear);

            multiReplaceStringRequest.InputFileBytes = formTemplateInput;
            return await editDocumentApi.EditDocumentDocxReplaceMultiAsync(multiReplaceStringRequest);

        }

        public byte[] MergeDocx(Stream inputFile1,Stream inputFile2)
        {
            return new MergeDocumentApi().MergeDocumentDocx(inputFile1, inputFile2);
        }
       
        public byte[] CreateBlankDoc(string intialContent = null)
        {
            intialContent = string.IsNullOrWhiteSpace(intialContent) ? "" : intialContent;

            EditDocumentApi editDocumentApi = new EditDocumentApi();
            CreateBlankDocxResponse createBlankDocxResponse = editDocumentApi.EditDocumentDocxCreateBlankDocument(new CreateBlankDocxRequest(intialContent));
            FinishEditingRequest finishEditingRequest = new FinishEditingRequest(createBlankDocxResponse.EditedDocumentURL);
            
            return editDocumentApi.EditDocumentFinishEditing(finishEditingRequest);
        }

        public byte[] CreatePDFCoverPage()
        {
            byte[] blankDocument = CreateBlankDoc("Title:Experimental Use Defense to Patent Infringement");
            
            return new ConvertDocumentApi()
                .ConvertDocumentDocxToPdf(new MemoryStream(blankDocument));
        }

        public byte[] CreatePDFCoverPageAsync()
        {
            HtmlToPdfRequest htmlToPdfRequest = new HtmlToPdfRequest("Title:Experimental Use Defense to Patent Infringement");

            return new ConvertWebApi().ConvertWebHtmlToPdf(htmlToPdfRequest);
        }

        public byte[] MergePdfs(Stream inputFile1, Stream inputFile2)
        {
            MergeDocumentApi mergeDocumentApi = new MergeDocumentApi();
            return mergeDocumentApi.MergeDocumentPdf(inputFile1, inputFile2);
        }

        public Task<byte[]> AddPdfAnnotations(byte[] inputData, List<Annotation> annotationsToAdd)
        {
            List<PdfAnnotation> annotations = new List<PdfAnnotation>();
            
            foreach (var annotation in annotationsToAdd)
            {
                annotations.Add(new PdfAnnotation(
                    annotation.Title, annotation.AnnotationType,
                    annotation.PageNumber, annotation.AnnotationIndex,
                    annotation.Subject, annotation.TextContents, annotation.CreationDate, 
                    annotation.ModifiedDate, annotation.LeftX, annotation.TopY,
                    annotation.Width, annotation.Height
                ));
            }

            AddPdfAnnotationRequest addPdfAnnotationRequest = new AddPdfAnnotationRequest(inputData, annotations);
            
            return new EditPdfApi()
                .EditPdfAddAnnotationsAsync(addPdfAnnotationRequest);
        }
        public IList<Annotation> GetPdfAnnotations(Stream inputFile)
        {
            EditPdfApi editPdfApi = new EditPdfApi();
            GetPdfAnnotationsResult result= editPdfApi.EditPdfGetAnnotations(inputFile);
            List<Annotation> annotations = new List<Annotation>();
            foreach (var annotation in result.Annotations)
            {
                annotations.Add(new Annotation {
                    Title = annotation.Title,
                    Subject = annotation.Subject,
                    AnnotationIndex = annotation.AnnotationIndex,
                    AnnotationType = annotation.AnnotationType,
                    ModifiedDate = annotation.ModifiedDate,
                    CreationDate = DateTime.Now,
                    TextContents = annotation.TextContents,
                    Width = annotation.Width,
                    Height = annotation.Height,
                    TopY = annotation.TopY,
                    LeftX = annotation.LeftX,
                    PageNumber = annotation.PageNumber

                });
            }
            return annotations;

        }
    }
}

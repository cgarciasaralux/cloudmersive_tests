
using DocumentConverter.Lib.Models;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace DocumentConverter.Lib
{
    public interface IDocumentManager
    {

        Task<byte[]> CreateFormCoverPage(FormCitation formCitation);
        Task<byte[]> CreateFormCoverPageHtml(FormCitation formCitation);
        byte[] MergeDocx(Stream inputFile1, Stream inputFile2);
        byte[] CreateBlankDoc(string initialContent=null);
        byte[] CreatePDFCoverPage();
        byte[] MergePdfs(Stream inputFile1, Stream inputFile2);
        Task<byte[]> AddPdfAnnotations(byte[] inputData, List<Annotation> annotationsToAdd);
        IList<Annotation> GetPdfAnnotations(Stream inputFile);
        Task<byte[]> CreateFormCoverPageTemplate(FormCitation formCitation,byte[] templateFile);
        
        //NEW
        Task<byte[]> CreateFormCoverPageTemplateAsync(FormCitation formCitation, byte[] formTemplateInput);
        Task<byte[]> CreateFormCoverPageHtmlAsync(FormCitation formCitation, string template);
        Task<byte[]> CreateFormCoverPageAsync(FormCitation formCitation, string template);
        byte[] CreatePDFCoverPageAsync();

    }
}
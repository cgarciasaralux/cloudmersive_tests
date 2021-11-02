using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Model;
using System.IO;

namespace DocumentConverter.Lib
{
    public interface IDocumentConverter
    {
        byte[] ConvertDocToPdf(Stream inputFile);
        byte[] ConvertPdfToDocx(Stream inputFile);
        byte[] ConvertRtfToDocx(Stream inputFile);
        byte[] ConvertHtmlToPdf(Stream inputFile);
        byte[] ConvertDocumentPptxToPdf(Stream inputFile);
        PptxToPngResult ConvertDocumentPptxToPng(Stream inputFile);
    }
}
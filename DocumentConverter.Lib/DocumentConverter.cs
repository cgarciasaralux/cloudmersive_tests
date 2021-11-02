using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Api;
using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Client;
using Cloudmersive.APIClient.NET.DocumentAndDataConvert.Model;
using System.IO;

namespace DocumentConverter.Lib
{
    public class DocumentConverter:IDocumentConverter
    {
        readonly ConvertDocumentApi convertDocumentApi;

        public DocumentConverter(string cloudMersiveApiKey)
        {
            Configuration.Default.AddApiKey("Apikey", cloudMersiveApiKey);
            convertDocumentApi = new ConvertDocumentApi();
        }

        public byte[] ConvertRtfToDocx(Stream stream)
        {
            return convertDocumentApi.ConvertDocumentRtfToDocx(stream);
        }

        public byte[] ConvertDocToPdf(Stream stream)
        {
            return convertDocumentApi.ConvertDocumentDocToPdf(stream);
        }

        public byte[] ConvertPdfToDocx(Stream stream)
        {
            return convertDocumentApi.ConvertDocumentPdfToDocx(stream);
        }
        public byte[] ConvertHtmlToPdf(Stream stream)
        {
            return convertDocumentApi.ConvertDocumentHtmlToPdf(stream);
        }
        public byte[] ConvertDocumentPptxToPdf(Stream inputFile)
        {
            return convertDocumentApi.ConvertDocumentPptxToPdf(inputFile);
        }
        public PptxToPngResult ConvertDocumentPptxToPng(Stream inputFile)
        {
            return convertDocumentApi.ConvertDocumentPptxToPng(inputFile);
        }
    }
}

using Accusoft.PrizmDocServer.Conversion;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Accusoft.PrizmDoc
{
    public interface IDocConverter
    {
        Task<IReadOnlyCollection<ConversionResult>> ConvertAsync(ConversionSourceDocument sourceDocument, DestinationFileFormat destinationFileFormat);
        Task ConvertRtfToDocxAsync(string sourceFile, string destinationFile);
    }
}
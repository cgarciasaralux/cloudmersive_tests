using Accusoft.PrizmDocServer;
using Accusoft.PrizmDocServer.Conversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Accusoft.PrizmDoc
{
    public class PrizmDocConverter : IDocConverter
    {
        private readonly PrizmDocServerClient _prizmDocServer;
        //MjXoHvf0ia9tjXOpFUgCny2XWEYOeQIOiLb3MIY0Nqu4XA7qzGxLm7jqSS8ykQPa
        public PrizmDocConverter(string apiKey)
        {
            _prizmDocServer = new PrizmDocServerClient("https://api.accusoft.com", apiKey);
        }
        public async Task<IReadOnlyCollection<ConversionResult>> ConvertAsync(ConversionSourceDocument sourceDocument, DestinationFileFormat destinationFileFormat)
        {
            var conversionResult = await _prizmDocServer.ConvertAsync(sourceDocument, destinationFileFormat);

            //conversionResult.Single().RemoteWorkFile.SaveAsync()


            return conversionResult;
        }
        public async Task ConvertRtfToDocxAsync(string sourceFile, string destinationFile)
        {
             var conversionResult = await _prizmDocServer.ConvertAsync(sourceFile, DestinationFileFormat.Pdf);
             await conversionResult.Single().RemoteWorkFile.SaveAsync(destinationFile);

        }
    }
}

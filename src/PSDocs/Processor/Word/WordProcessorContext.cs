using PSDocs.Configuration;
using PSDocs.Models;

namespace PSDocs.Processor.Word
{
    internal sealed class WordProcessorContext
    {
        public readonly PSDocumentOption Option;
        public readonly Document Document;
        public readonly OpenXmlBuilder Builder;

        public WordProcessorContext(PSDocumentOption option, Document document, string templateFile)
        {
            Option = option;
            Document = document;

            Builder = new OpenXmlBuilder(templateFile);
        }
    }
}
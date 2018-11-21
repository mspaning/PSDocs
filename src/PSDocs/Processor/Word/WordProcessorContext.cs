using DocumentFormat.OpenXml.Packaging;
using PSDocs.Configuration;
using PSDocs.Models;

namespace PSDocs.Processor.Word
{
    internal class WordProcessorContext
    {
        public readonly PSDocumentOption Option;
        public readonly Document Document;
        public readonly DocxBuilder Builder;

        public WordProcessorContext(PSDocumentOption option, Document document, string templateFile)
        {
            Option = option;
            Document = document;

            Builder = new DocxBuilder(templateFile);
        }
    }
}
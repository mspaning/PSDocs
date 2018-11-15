using DocumentFormat.OpenXml.Packaging;
using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace PSDocs.Processor.Word
{
    public sealed class DocxBuilder
    {
        private WordprocessingDocument _Document;

        public DocxBuilder()
        {

            //_Docx = new 
        }

        public OpenXml.Body Body
        {
            get
            {
                return _Document.MainDocumentPart.Document.Body;
            }
        }

        public void CreateDocument(string filepath)
        {
            // Create a document at a file path
            _Document = WordprocessingDocument.Create(filepath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);

            // Add a main document part. 
            var mainPart = _Document.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new OpenXml.Document();
            var body = mainPart.Document.AppendChild(new OpenXml.Body());
            var para = body.AppendChild(new OpenXml.Paragraph());
            var run = para.AppendChild(new OpenXml.Run());
            run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
        }

        public void Build()
        {
            _Document.Save();
        }
    }
}
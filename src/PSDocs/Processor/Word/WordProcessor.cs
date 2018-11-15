using PSDocs.Configuration;
using PSDocs.Models;
using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace PSDocs.Processor.Word
{
    public sealed class WordProcessor
    {

        public string Process(PSDocumentOption option, Document document)
        {
            if (document == null)
            {
                return string.Empty;
            }

            var context = new WordProcessorContext(option, document);
            context.Builder.CreateDocument("c:\\dev\\workspace\\PSDocs\\demo.docx");

            Document(context);

            context.Builder.Build();

            //context.Builder.Remove(context.Builder.Length - 2, 2);
            //var result = context.Builder.ToString();

            //if (string.IsNullOrEmpty(result))
            //{
            //    result = null;
            //}

            //return result;

            return null;
        }

        private void Document(WordProcessorContext context)
        {
            // Process metadata
            //Metadata(context);

            //if (!string.IsNullOrEmpty(context.Document.Title))
            //{
            //    context.WriteLine("# ", context.Document.Title);
            //    context.WriteLine(string.Empty);
            //}

            foreach (var node in context.Document.Node)
            {
                Node(context, node);
            }
        }

        private void Node(WordProcessorContext context, object node)
        {
            if (node == null)
            {
                return;
            }

            var documentNode = node as DocumentNode;

            if (documentNode != null)
            {
                switch (documentNode.Type)
                {
                    case DocumentNodeType.Section:

                        Section(context, documentNode as Section);

                        break;

                    //case DocumentNodeType.Table:

                    //    Table(context, documentNode as Table);

                    //    break;

                    //case DocumentNodeType.Code:

                    //    Code(context, documentNode as Code);

                    //    break;

                    //case DocumentNodeType.BlockQuote:

                    //    BlockQuote(context, documentNode as BlockQuote);

                    //    break;

                    //case DocumentNodeType.Text:

                    //    Text(context, documentNode as Text);

                    //    break;

                    //case DocumentNodeType.Include:

                    //    Include(context, documentNode as Include);

                    //    break;
                }

                return;
            }

            //String(context, node.ToString());
        }

        private void Section(WordProcessorContext context, Section section)
        {
            var body = context.Builder.Body;
            var paragraph = body.AppendChild(new OpenXml.Paragraph());
            var run = paragraph.AppendChild(new OpenXml.Run());
            run.AppendChild(new OpenXml.Text(section.Title));

            if (section.Node.Count > 0)
            {
                foreach (var node in section.Node)
                {
                    Node(context, node);
                }
            }
        }
    }
}

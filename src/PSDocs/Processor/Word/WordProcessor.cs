using PSDocs.Configuration;
using PSDocs.Models;
using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace PSDocs.Processor.Word
{
    public sealed class WordProcessor
    {

        public string Process(PSDocsContext executionContext, Document document)
        {
            if (document == null)
            {
                return string.Empty;
            }

            var templatePath = System.IO.Path.Combine(executionContext.ModulePath, "Templates/Template.docx");

            var context = new WordProcessorContext(executionContext.Option, document, templatePath);
            context.Builder.Document(System.IO.Path.Combine(document.Path, string.Concat(document.DocumentName, ".docx")));

            Document(context);

            context.Builder.Build();

            return null;
        }

        private void Document(WordProcessorContext context)
        {
            // TODO: Handle metadata
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

                    case DocumentNodeType.Table:

                        Table(context, documentNode as Table);

                        break;

                    case DocumentNodeType.Code:

                        Code(context, documentNode as Code);

                        break;

                    case DocumentNodeType.BlockQuote:

                        BlockQuote(context, documentNode as BlockQuote);

                        break;

                    case DocumentNodeType.Text:

                        Text(context, documentNode as Text);

                        break;

                        // TODO: Handle includes
                        //case DocumentNodeType.Include:

                        //    Include(context, documentNode as Include);

                        //    break;
                }

                return;
            }

            //String(context, node.ToString());
        }

        private void Section(WordProcessorContext context, Section node)
        {
            context.Builder.Section(node.Title);

            if (node.Node.Count > 0)
            {
                foreach (var child in node.Node)
                {
                    Node(context, child);
                }
            }
        }

        private void BlockQuote(WordProcessorContext context, BlockQuote node)
        {
            context.Builder.BlockQuote(node.Content);
        }

        private void Code(WordProcessorContext context, Code node)
        {
            context.Builder.Code(node.Content);
        }

        private void Text(WordProcessorContext context, Text node)
        {
            context.Builder.Text(node.Content);
        }

        private void Table(WordProcessorContext context, Table node)
        {
            context.Builder.Table(node.Rows);
        }
    }
}

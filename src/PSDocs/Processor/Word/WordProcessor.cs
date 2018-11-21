using PSDocs.Models;

namespace PSDocs.Processor.Word
{
    /// <summary>
    /// A processor to handle document generation for OpenXml (.docx).
    /// </summary>
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

            // Handle the node
            if (node is DocumentNode documentNode)
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

                    case DocumentNodeType.Include:

                        Include(context, documentNode as Include);

                        break;
                }

                return;
            }

            //String(context, node.ToString());
        }

        /// <summary>
        /// Process sections.
        /// </summary>
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

        /// <summary>
        /// Process block quotes.
        /// </summary>
        private void BlockQuote(WordProcessorContext context, BlockQuote node)
        {
            context.Builder.BlockQuote(node.Content);
        }

        /// <summary>
        /// Process code blocks.
        /// </summary>
        private void Code(WordProcessorContext context, Code node)
        {
            context.Builder.Code(node.Content);
        }

        /// <summary>
        /// Process text.
        /// </summary>
        private void Text(WordProcessorContext context, Text node)
        {
            context.Builder.Text(node.Content);
        }

        /// <summary>
        /// Process tables.
        /// </summary>
        private void Table(WordProcessorContext context, Table node)
        {
            context.Builder.Table(node.Rows);
        }

        /// <summary>
        /// Process includes.
        /// </summary>
        private void Include(WordProcessorContext context, Include node)
        {
            var text = System.IO.File.ReadAllText(node.Path);
            context.Builder.Text(text);
        }
    }
}

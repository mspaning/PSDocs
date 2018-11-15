using PSDocs.Configuration;
using PSDocs.Models;
using PSDocs.Processor.Word;
using Xunit;

namespace PSDocs
{
    public class OpenXmlProcessorTests
    {
        [Fact]
        public void Test1()
        {
            var processor = GetProcessor();
            var options = new PSDocumentOption();


            processor.Process(option: options, document: GetDocument());
        }

        private static WordProcessor GetProcessor()
        {
            return new WordProcessor();
        }

        private static Document GetDocument()
        {
            var result = ModelHelper.NewDocument();
            result.Title = "OpenXmlProcessorTests";

            result.Node.Add(ModelHelper.NewSection("Section 1", 2));
            result.Node.Add(ModelHelper.Text("This is section 1."));

            return result;
        }
    }
}

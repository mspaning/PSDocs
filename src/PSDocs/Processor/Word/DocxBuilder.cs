using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Linq;
using OpenXml = DocumentFormat.OpenXml.Wordprocessing;

namespace PSDocs.Processor.Word
{
    public sealed class DocxBuilder
    {
        private WordprocessingDocument _Document;
        private readonly string _TemplateFile;

        public DocxBuilder(string templateFile)
        {
            _TemplateFile = templateFile;
            //_Docx = new 
        }

        public OpenXml.Body Body
        {
            get
            {
                return _Document.MainDocumentPart.Document.Body;
            }
        }

        public void Document(string filePath)
        {
            // Create a document at a file path
            System.IO.File.Copy(_TemplateFile, filePath, true);
            _Document = WordprocessingDocument.Open(filePath, true);

            if (_Document.MainDocumentPart == null)
            {
                // Add a main document part. 
                var mainPart = _Document.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new OpenXml.Document();
                
            }

            // Create the document structure and add some text.
            _Document.MainDocumentPart.Document = new OpenXml.Document();

            var body = _Document.MainDocumentPart.Document.Body = new OpenXml.Body();
                //.AppendChild(new OpenXml.Body());
            //var para = body.AppendChild(new OpenXml.Paragraph());
            //var run = para.AppendChild(new OpenXml.Run());
            //run.AppendChild(new OpenXml.Text("Create text in body - CreateWordprocessingDocument"));

            _Document.Save();

            // Get the Styles part for this document.
            var styles = _Document.MainDocumentPart.StyleDefinitionsPart;

            if (styles == null)
            {
                styles = _Document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                styles.Styles = new OpenXml.Styles();
            }
        }

        public void Section(string title)
        {
            var paragraph = Body.AppendChild(new OpenXml.Paragraph());
            var properties = paragraph.AppendChild(new OpenXml.ParagraphProperties());
            var run = paragraph.AppendChild(new OpenXml.Run());
            run.AppendChild(new OpenXml.Text(title));

            // Set the style of the paragraph.
            properties.ParagraphStyleId = new OpenXml.ParagraphStyleId
            {
                Val = GetStyleId("heading 2")
            };
        }

        public void Text(string text)
        {
            var para = Body.AppendChild(new OpenXml.Paragraph());
            var run = para.AppendChild(new OpenXml.Run());
            run.AppendChild(new OpenXml.Text(text));
        }

        public void BlockQuote(string[] text)
        {
            var paragraph = Body.AppendChild(new OpenXml.Paragraph());
            var properties = paragraph.AppendChild(new OpenXml.ParagraphProperties());

            foreach (var s in text)
            {
                var run = paragraph.AppendChild(new OpenXml.Run());
                run.AppendChild(new OpenXml.Text(s));
            }

            // Set the style of the paragraph.
            properties.ParagraphStyleId = new OpenXml.ParagraphStyleId
            {
                Val = GetStyleId("block quote")
            };
        }

        public void Code(string text)
        {
            var paragraph = Body.AppendChild(new OpenXml.Paragraph());
            var properties = paragraph.AppendChild(new OpenXml.ParagraphProperties());
            var run = paragraph.AppendChild(new OpenXml.Run());
            run.AppendChild(new OpenXml.Text(text));

            // Set the style of the paragraph.
            properties.ParagraphStyleId = new OpenXml.ParagraphStyleId
            {
                Val = GetStyleId("code")
            };
        }

        // Take the data from a two-dimensional array and build a table at the 
        // end of the supplied document.
        public void Table(List<string[]> data)
        {
            var document = _Document.MainDocumentPart.Document;

            var table = new OpenXml.Table();

            var properties = new OpenXml.TableProperties(
                new OpenXml.TableBorders(
                new OpenXml.TopBorder
                {
                    Val = new DocumentFormat.OpenXml.EnumValue<OpenXml.BorderValues>(OpenXml.BorderValues.Single),
                    Size = 12
                },
                new OpenXml.BottomBorder
                {
                    Val = new DocumentFormat.OpenXml.EnumValue<OpenXml.BorderValues>(OpenXml.BorderValues.Single),
                    Size = 12
                },
                new OpenXml.LeftBorder
                {
                    Val = new DocumentFormat.OpenXml.EnumValue<OpenXml.BorderValues>(OpenXml.BorderValues.Single),
                    Size = 12
                },
                new OpenXml.RightBorder
                {
                    Val = new DocumentFormat.OpenXml.EnumValue<OpenXml.BorderValues>(OpenXml.BorderValues.Single),
                    Size = 12
                },
                new OpenXml.InsideHorizontalBorder
                {
                    Val = new DocumentFormat.OpenXml.EnumValue<OpenXml.BorderValues>(OpenXml.BorderValues.Single),
                    Size = 12
                },
                new OpenXml.InsideVerticalBorder
                {
                    Val = new DocumentFormat.OpenXml.EnumValue<OpenXml.BorderValues>(OpenXml.BorderValues.Single),
                    Size = 12
                }));

            table.AppendChild<OpenXml.TableProperties>(properties);

            for (var i = 0; i < data.Count; i++)
            {
                var tr = new OpenXml.TableRow();
                for (var j = 0; j < data[i].Length; j++)
                {
                    var tc = new OpenXml.TableCell();
                    tc.Append(new OpenXml.Paragraph(new OpenXml.Run(new OpenXml.Text(data[i][j]))));

                    // Assume you want columns that are automatically sized.
                    tc.Append(new OpenXml.TableCellProperties(
                        new OpenXml.TableCellWidth { Type = OpenXml.TableWidthUnitValues.Auto }));

                    tr.Append(tc);
                }
                table.Append(tr);
            }
            document.Body.Append(table);
            document.Save();

        }

        private void SetStyle(string styleName)
        {
            // Get the Styles part for this document.
            var styles = _Document.MainDocumentPart.StyleDefinitionsPart;

            if (styles == null)
            {
                styles = _Document.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                styles.Styles = new OpenXml.Styles();
            }

            // No match on styleid, so let's try style name.
            var styleId = GetStyleId(styleName);

            if (styleId == null)
            {
                //AddNewStyle(styles, styleId, styleName);
            }
        }

        public string GetStyleId(string styleName)
        {
            var styles = _Document.MainDocumentPart.StyleDefinitionsPart.Styles;

            var styleId = styles
                .Descendants<OpenXml.StyleName>()
                .Where(s => s.Val.Value.Equals(styleName) && (((OpenXml.Style)s.Parent).Type == OpenXml.StyleValues.Paragraph))
                .Select(n => ((OpenXml.Style)n.Parent).StyleId)
                .FirstOrDefault();

            return styleId;
        }

        public void Build()
        {
            _Document.Dispose();
            //_Document.Save();
        }

        public void Validate()
        {
            try
            {
                OpenXmlValidator validator = new OpenXmlValidator();
                int count = 0;
                foreach (ValidationErrorInfo error in validator.Validate(_Document))
                {
                    count++;

                    throw new System.Exception(error.Description);
                    //Console.WriteLine("Error " + count);
                    //Console.WriteLine("Description: " + error.Description);
                    //Console.WriteLine("ErrorType: " + error.ErrorType);
                    //Console.WriteLine("Node: " + error.Node);
                    //Console.WriteLine("Path: " + error.Path.XPath);
                    //Console.WriteLine("Part: " + error.Part.Uri);
                    //Console.WriteLine("-------------------------------------------");
                }
            }
            finally
            {

            }
        }
    }
}
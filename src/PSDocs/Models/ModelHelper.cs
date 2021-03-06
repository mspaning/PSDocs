﻿
using PSDocs.Configuration;
using PSDocs.Resources;
using System.IO;

namespace PSDocs.Models
{
    internal static class ModelHelper
    {
        public static Document NewDocument()
        {
            return new Document(null, null);
        }

        public static Section NewSection(string name, int level)
        {
            return new Section
            {
                Title = name,
                Level = level
            };
        }

        public static TableBuilder Table()
        {
            return new TableBuilder();
        }

        public static Code NewCode()
        {
            return new Code
            {

            };
        }

        public static BlockQuote BlockQuote(string info, string title)
        {
            return new BlockQuote
            {
                Info = info,
                Title = title
            };
        }

        public static Text Text(string value)
        {
            return new Text
            {
                Content = value
            };
        }

        public static Include Include(string baseDirectory, string culture, string fileName, bool useCulture)
        {
            baseDirectory = PSDocumentOption.GetRootedPath(baseDirectory);
            var absolutePath = Path.IsPathRooted(fileName) ? fileName : Path.Combine(baseDirectory, (useCulture ? culture : string.Empty), fileName);
            if (!File.Exists(absolutePath))
                throw new FileNotFoundException(PSDocsResources.IncludeNotFound, absolutePath);

            var text = File.ReadAllText(absolutePath);
            return new Include
            {
                Path = absolutePath,
                Content = text,
            };
        }
    }
}

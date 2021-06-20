﻿namespace MariGold.OpenXHTML
{
    using System;
    using System.Collections.Generic;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.IO;
    using MariGold.HtmlParser;

    public sealed class WordDocument : IDisposable
    {
        private readonly IOpenXmlContext context;

        public string ImagePath
        {
            get
            {
                return context.ImagePath;
            }

            set
            {
                context.ImagePath = value;
            }
        }

        public string BaseURL
        {
            get
            {
                return context.BaseURL;
            }

            set
            {
                context.BaseURL = value;
            }
        }

        public string UriSchema
        {
            get
            {
                return context.UriSchema;
            }

            set
            {
                context.UriSchema = value;
            }
        }

        public WordprocessingDocument WordprocessingDocument
        {
            get
            {
                return context.WordprocessingDocument;
            }
        }

        public MainDocumentPart MainDocumentPart
        {
            get
            {
                return context.MainDocumentPart;
            }
        }

        public Document Document
        {
            get
            {
                return context.Document;
            }
        }

        public WordDocument(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentNullException("fileName");
            }

            context = new OpenXmlContext(WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document));
        }

        public WordDocument(Stream stream)
        {
            if (stream == null)
            {
                throw new ArgumentNullException("stream");
            }

            context = new OpenXmlContext(WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document));
        }

        public void Process(IParser parser)
        {
            if (parser == null)
            {
                throw new ArgumentNullException("parser");
            }

            parser.BaseURL = context.BaseURL;
            parser.UriSchema = context.UriSchema;
            IHtmlNode node = parser.FindBodyOrFirstElement();
            context.SetParser(parser);

            if (node != null)
            {
                DocxElement body = context.GetBodyElement();
                Paragraph paragraph = null;
                body.Process(new DocxNode(node), ref paragraph, new Dictionary<string, object>());
            }
        }

        public void Save()
        {
            context.Save();
        }

        public void Close()
        {
            context.WordprocessingDocument.Close();
        }

        public void SaveAndClose()
        {
            Save();
            Close();
        }

        public void Dispose()
        {
            context.Dispose();
        }
    }
}

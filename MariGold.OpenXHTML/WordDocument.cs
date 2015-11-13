namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.IO;
	using MariGold.HtmlParser;
	
	/// <summary>
	/// 
	/// </summary>
	public sealed class WordDocument
	{
		private readonly IWordContext context;
		
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
			
			context = new WordContext(WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document));
		}
		
		public WordDocument(MemoryStream stream)
		{
			if (stream == null)
			{
				throw new ArgumentNullException("stream");
			}
			
			context = new WordContext(WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document));
		}
		
		public void Process(IParser parser)
		{
			if (parser == null)
			{
				throw new ArgumentNullException("parser");
			}
		
			HtmlNode node = parser.FindBodyOrFirstElement();
			
			while (node != null)
			{
				WordElement element = context.Convert(node);
				
				if (element != null)
				{
					element.Process(node);
				}
				
				node = node.Next;
			}
		}
		
		public void Save()
		{
			context.Clear();
		}
	}
}

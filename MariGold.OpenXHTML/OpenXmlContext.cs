namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.Collections.Generic;
	using MariGold.HtmlParser;
	
	internal sealed class OpenXmlContext : IOpenXmlContext
	{
		private WordprocessingDocument document;
		private MainDocumentPart mainPart;
		private List<WordElement> elements;
		
		private void PrepareWordElements()
		{
			elements = new List<WordElement>() 
			{
				new DocxDiv(this),
				new DocxSpan(this)
			};
		}
		
		internal OpenXmlContext(WordprocessingDocument document)
		{
			this.document = document;
			mainPart = this.document.AddMainDocumentPart();
			mainPart.Document = new Document();
			
			PrepareWordElements();
		}
		
		public WordprocessingDocument WordprocessingDocument
		{
			get
			{
				if (document == null)
				{
					throw new InvalidOperationException("Document is not opened!");
				}
				
				return document;
			}
		}
		
		public MainDocumentPart MainDocumentPart
		{
			get
			{
				if (mainPart == null)
				{
					throw new InvalidOperationException("Document is not opened!");
				}
				
				return mainPart;
			}
		}
		
		public Document Document
		{
			get
			{
				return MainDocumentPart.Document;
			}
		}
		
		public void Clear()
		{
			document.Close();
			document.Dispose();
			
			document = null;
			mainPart = null;
		}
		
		public WordElement Convert(HtmlNode node)
		{
			foreach (WordElement element in elements)
			{
				if (element.CanConvert(node))
				{
					return element;
				}
			}
			
			return null;
		}
		
		public WordElement GetBodyElement()
		{
			return new DocxBody(this);
		}
	}
}

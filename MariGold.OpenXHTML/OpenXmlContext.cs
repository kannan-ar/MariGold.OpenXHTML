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
		private List<DocxElement> elements;
		
		private void PrepareWordElements()
		{
			elements = new List<DocxElement>()
			{
				new DocxDiv(this),
				new DocxUL(this),
				new DocxImage(this),
				new DocxSpan(this),
				new DocxA(this),
				new DocxBr(this),
				new DocxUnderline(this),
				new DocxItalic(this),
				new DocxBold(this),
				new DocxHeader(this),
				new DocxTable(this)
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
			Document.Save();
			
			document.Close();
			document.Dispose();
			
			document = null;
			mainPart = null;
		}
		
		public DocxElement Convert(IHtmlNode node)
		{
			foreach (DocxElement element in elements)
			{
				if (element.CanConvert(node))
				{
					return element;
				}
			}
			
			return null;
		}
		
		public DocxElement GetBodyElement()
		{
			return new DocxBody(this);
		}
	}
}

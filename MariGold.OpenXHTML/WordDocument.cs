namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.IO;
	
	/// <summary>
	/// 
	/// </summary>
	public sealed class WordDocument
	{
		private WordprocessingDocument document;
		private MainDocumentPart mainPart;
		
		private void PrepareWordDocument(WordprocessingDocument document)
		{
			this.document = document;
			mainPart = this.document.AddMainDocumentPart();
			mainPart.Document = new Document();
		}
		
		private void Clear()
		{
			document = null;
			mainPart = null;
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
		
		public WordDocument(string fileName)
		{
			PrepareWordDocument(WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document));
		}
		
		public WordDocument(MemoryStream stream)
		{
			PrepareWordDocument(WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document));
		}
		
		public void Process(IParser parser)
		{
			throw new NotImplementedException();
		}
		
		public void Save()
		{
			document.Close();
			document.Dispose();
			
			Clear();
		}
	}
}

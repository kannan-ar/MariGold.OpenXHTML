namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.Collections.Generic;
	using MariGold.HtmlParser;
	using System.Linq;
	
	internal sealed class OpenXmlContext : IOpenXmlContext
	{
		private WordprocessingDocument document;
		private MainDocumentPart mainPart;
		private List<DocxElement> elements;
		private Dictionary<NumberFormatValues,AbstractNum> abstractNumList;
		private Dictionary<NumberFormatValues,NumberingInstance> numberingInstanceList;
		
		private void PrepareWordElements()
		{
			elements = new List<DocxElement>()
			{
				new DocxDiv(this),
				new DocxUL(this),
				new DocxOL(this),
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
		
		private void SaveNumberDefinitions()
		{
			if (abstractNumList != null && numberingInstanceList != null)
			{
				if (mainPart.NumberingDefinitionsPart == null)
				{
					NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>("numberingDefinitionsPart");
				}
			
				Numbering numbering = new Numbering();
			
				foreach (var abstractNum in abstractNumList)
				{
					numbering.Append(abstractNum.Value);
				}
			
				foreach (var numberingInstance in numberingInstanceList)
				{
					numbering.Append(numberingInstance.Value);
				}
			
				mainPart.NumberingDefinitionsPart.Numbering = numbering;
			}
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
		
		public void Save()
		{
			SaveNumberDefinitions();
			
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
		
		public bool HasNumberingDefinition(NumberFormatValues format)
		{
			return abstractNumList != null && numberingInstanceList != null && abstractNumList.ContainsKey(format) && numberingInstanceList.ContainsKey(format);
		}
		
		public void SaveNumberingDefinition(NumberFormatValues format, AbstractNum abstractNum, NumberingInstance numberingInstance)
		{
			if (abstractNumList == null)
			{
				abstractNumList = new Dictionary<NumberFormatValues, AbstractNum>();
			}
			
			if (numberingInstanceList == null)
			{
				numberingInstanceList = new Dictionary<NumberFormatValues, NumberingInstance>();
			}
			
			if (!abstractNumList.ContainsKey(format))
			{
				abstractNumList.Add(format, abstractNum);
			}
			
			if (!numberingInstanceList.ContainsKey(format))
			{
				numberingInstanceList.Add(format, numberingInstance);
			}
		}
	}
}

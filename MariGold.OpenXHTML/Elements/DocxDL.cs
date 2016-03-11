namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxDL : DocxElement
	{
		private const string defaultDDLeftMargin = "40px";
		private const string defaultDLMargin = "1em";
		
		private void ProcessChild(IHtmlNode node, OpenXmlElement parent)
		{
			if (node == null)
			{
				return;
			}
			
			Paragraph paragraph = parent.AppendChild(new Paragraph());
			ParagraphCreated(node, paragraph);
			
			foreach (IHtmlNode child in node.Children)
			{
				if (child.IsText)
				{
					Run run = paragraph.AppendChild(new Run(new Text(child.InnerHtml)));
					RunCreated(node, run);
				}
				else
				{
					ProcessChild(child, parent, ref paragraph);
				}
			}
		}
		
		private void SetDDProperties(IHtmlNode node)
		{
			DocxMargin margin = new DocxMargin(node);
			
			string leftMargin = margin.GetLeftMargin();
			
			if (string.IsNullOrEmpty(leftMargin))
			{
				//Default left margin of dd element
				margin.SetLeftMargin(defaultDDLeftMargin);
			}
		}
		
		private void SetMarginTop(OpenXmlElement parent)
		{
			Paragraph para = parent.AppendChild(new Paragraph());
			para.ParagraphProperties = new ParagraphProperties();
			
			DocxMargin.SetTopMargin(defaultDLMargin, para.ParagraphProperties);
		}
		
		private void SetMarginBottom(OpenXmlElement parent)
		{
			Paragraph para = parent.AppendChild(new Paragraph());
			para.ParagraphProperties = new ParagraphProperties();
			
			DocxMargin.SetBottomMargin(defaultDLMargin, para.ParagraphProperties);
		}
		
		internal DocxDL(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "dl", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node == null || parent == null || !CanConvert(node))
			{
				return;
			}
			
			if (!node.HasChildren)
			{
				return;
			}
			
			paragraph = null;
			
			//Add an empty paragraph to set default margin top
			SetMarginTop(parent);
			
			foreach (IHtmlNode child in node.Children)
			{
				if (string.Compare(child.Tag, "dt", StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					ProcessChild(child, parent);
				}
				else
				if (string.Compare(child.Tag, "dd", StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					SetDDProperties(child);
					ProcessChild(child, parent);
				}
			}
			
			//Add an empty paragraph at the end to set default margin bottom
			SetMarginBottom(parent);
		}
	}
}

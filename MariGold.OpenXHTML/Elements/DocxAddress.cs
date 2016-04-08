namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxAddress : DocxElement
	{
		private void SetDefaultStyle(IHtmlNode node)
		{
			DocxNode docxNode = new DocxNode(node);
			
			string value = docxNode.ExtractStyleValue(DocxFont.fontStyle);
			
			if (string.IsNullOrEmpty(value))
			{
				docxNode.SetStyleValue(DocxFont.fontStyle, DocxFont.italic);
			}
		}
		
		internal DocxAddress(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "address", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node == null || parent == null)
			{
				return;
			}
			
			//Address tag also creats a new block element. Thus clear the existing paragraph
			paragraph = null;
			Paragraph addrParagraph = null;
			
			foreach (IHtmlNode child in node.Children)
			{
				SetDefaultStyle(child);
				
				if (child.IsText && !IsEmptyText(child.InnerHtml))
				{
					if (addrParagraph == null)
					{
						addrParagraph = parent.AppendChild(new Paragraph());
						ParagraphCreated(node, addrParagraph);
					}
						
					Run run = addrParagraph.AppendChild(new Run(new Text() {
						Text = ClearHtml(child.InnerHtml),
						Space = SpaceProcessingModeValues.Preserve
					}));
					RunCreated(child, run);
				}
				else
				{
					//Child elements will create on new address paragraph
					ProcessChild(child, parent, ref addrParagraph);
				}
			}
		}
	}
}

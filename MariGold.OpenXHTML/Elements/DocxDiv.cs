namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxDiv : DocxElement
	{
		internal DocxDiv(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "div", StringComparison.InvariantCultureIgnoreCase) == 0 ||
			string.Compare(node.Tag, "p", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node != null && parent != null)
			{
				//Div creates it's own new paragraph. So old paragraph ends here and creats another one after this div 
				//if there any text!
				paragraph = null;
				Paragraph divParagraph = null;
				
				foreach (IHtmlNode child in node.Children)
				{
					if (child.IsText && !IsEmptyText(child.InnerHtml))
					{
						if (divParagraph == null)
						{
							divParagraph = parent.AppendChild(new Paragraph());
							ParagraphCreated(node, divParagraph);
						}
						
						Run run = divParagraph.AppendChild(new Run(new Text() {
							Text = ClearHtml(child.InnerHtml),
							Space = SpaceProcessingModeValues.Preserve
						}));
						
						RunCreated(child, run);
					}
					else
					{
						//ProcessChild forwards the incomming parent to the child element. So any div element inside this div
						//creates a new paragraph on the parent element.
						ProcessChild(child, parent, ref divParagraph);
					}
				}
			}
		}
	}
}

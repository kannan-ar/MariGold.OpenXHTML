namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxCenter : DocxElement
	{
		internal DocxCenter(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "center", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node == null)
			{
				return;
			}
			
			foreach (IHtmlNode child in node.Children)
			{
				if (child.IsText)
				{
					if (!IsEmptyText(child.InnerHtml))
					{
						if (paragraph == null)
						{
							paragraph = parent.AppendChild(new Paragraph());
							IHtmlNode parentNode = node.Parent ?? node;
						
							ParagraphCreated(parentNode, paragraph);
						}
					
						if (paragraph.ParagraphProperties == null)
						{
							paragraph.ParagraphProperties = new ParagraphProperties();
						}

						DocxAlignment.AlignCenter(paragraph.ParagraphProperties);
					
						Run run = paragraph.AppendChild(new Run());
						RunCreated(node, run);
					
						run.AppendChild(new Text() {
							Text = ClearHtml(child.InnerHtml),
							Space = SpaceProcessingModeValues.Preserve
						});
					}
				}
				else
				{
					ProcessChild(child, parent, ref paragraph);
				}
			}
		}
	}
}

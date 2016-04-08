namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxUnderline : DocxElement
	{
		internal DocxUnderline(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "u", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node == null)
			{
				return;
			}
			
			foreach (IHtmlNode child in node.Children)
			{
				if (child.IsText && !IsEmptyText(child.InnerHtml))
				{
					if (paragraph == null)
					{
						paragraph = parent.AppendChild(new Paragraph());
						IHtmlNode parentNode = node.Parent ?? node;
						
						ParagraphCreated(parentNode, paragraph);
					}
					
					Run run = paragraph.AppendChild(new Run());
					RunCreated(node, run);
					
					if (run.RunProperties == null)
					{
						run.RunProperties = new RunProperties();
					}
					
					DocxFont.ApplyUnderline(run.RunProperties);
					
					run.AppendChild(new Text() {
						Text = ClearHtml(child.InnerHtml),
						Space = SpaceProcessingModeValues.Preserve
					});
				}
				else
				{
					ProcessChild(child, parent, ref paragraph);
				}
			}
		}
	}
}

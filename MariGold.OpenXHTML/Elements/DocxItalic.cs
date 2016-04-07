namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxItalic : DocxElement
	{
		internal DocxItalic(IOpenXmlContext context)
			:base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "i", StringComparison.InvariantCultureIgnoreCase) == 0 ||
				string.Compare(node.Tag, "em", StringComparison.InvariantCultureIgnoreCase) == 0;
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
						IHtmlNode parentNode = node.Parent??node;
						
						ParagraphCreated(parentNode, paragraph);
					}
					
					Run run = paragraph.AppendChild(new Run());
					RunCreated(node, run);
					
					if (run.RunProperties == null)
					{
						run.RunProperties = new RunProperties();
					}
					
					DocxFont.ApplyFontItalic(run.RunProperties);
					
					run.AppendChild(new Text(ClearHtml(child.InnerHtml)));
				}
				else
				{
					ProcessChild(child, parent, ref paragraph);
				}
			}
		}
	}
}

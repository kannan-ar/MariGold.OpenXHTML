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
				if (child.IsText)
				{
					if (paragraph == null)
					{
						paragraph = parent.AppendChild(new Paragraph());
						ParagraphCreated(node, paragraph);
					}
					
					Run run = paragraph.AppendChild(new Run());
					RunCreated(child, run);
					
					if (run.RunProperties == null)
					{
						run.RunProperties = new RunProperties();
					}
					
					DocxFont.ApplyUnderline(run.RunProperties);
					
					run.AppendChild(new Text(child.InnerHtml));
				}
				else
				{
					ProcessChild(child, parent, ref paragraph);
				}
			}
		}
	}
}

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
		
		internal override bool CanConvert(HtmlNode node)
		{
			return string.Compare(node.Tag, "i", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(HtmlNode node, OpenXmlElement parent)
		{
			if (node == null)
			{
				return;
			}
			
			foreach (HtmlNode child in node.Children)
			{
				if (child.IsText)
				{
					Run run = CreateRun(child);
					
					if (run.RunProperties == null)
					{
						run.RunProperties = new RunProperties();
					}
					
					DocxFont.ApplyFontItalic(run.RunProperties);
					
					run.AppendChild(new Text(child.InnerHtml));
					
					AppendToParagraph(node, parent, run);
				}
				else
				{
					ProcessChild(child, parent);
				}
			}
		}
	}
}

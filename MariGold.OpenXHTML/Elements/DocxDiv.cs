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
		
		internal override bool CanConvert(HtmlNode node)
		{
			return string.Compare(node.Tag, "div", true) == 0 ||
				string.Compare(node.Tag, "p", true) == 0;
		}
		
		internal override void Process(HtmlNode node, OpenXmlElement parent)
		{
			if (node != null && parent != null)
			{
				Parent.Current = null;
				OpenXmlElement paragraph = CreateParagraph(node, parent);
				
				foreach (HtmlNode child in node.Children)
				{
					if (child.IsText)
					{
						AppendRun(node, paragraph).AppendChild(new Text(node.InnerHtml));
					}
					else
					{
						ProcessChild(child, paragraph);
					}
				}
			}
		}
	}
}

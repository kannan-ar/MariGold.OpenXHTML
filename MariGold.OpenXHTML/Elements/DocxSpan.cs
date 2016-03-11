namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxSpan : DocxElement
	{
		public DocxSpan(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "span", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node != null && parent != null)
			{
				foreach (IHtmlNode child in node.Children)
				{
					if (child.IsText)
					{
						if (paragraph == null)
						{
							paragraph = parent.AppendChild(new Paragraph());
							IHtmlNode parentNode = node.Parent??node;
							
							ParagraphCreated(parentNode, paragraph);
						}
						
						Run run = paragraph.AppendChild(new Run(new Text(child.InnerHtml)));
						RunCreated(node, run);
					}
					else
					{
						ProcessChild(child, parent, ref paragraph);
					}
				}
			}
		}
	}
}

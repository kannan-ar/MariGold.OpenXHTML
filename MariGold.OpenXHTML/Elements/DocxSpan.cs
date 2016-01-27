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
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent)
		{
			if (node != null && parent != null)
			{
				Run	run = null;
				
				foreach (IHtmlNode child in node.Children)
				{
					if (child.IsText)
					{
						if (run == null)
						{
							run = AppendRun(node, parent);
						}
						
						run.AppendChild(new Text(child.InnerHtml));
					}
					else
					{
						ProcessChild(child, run);
					}
				}
			}
		}
	}
}

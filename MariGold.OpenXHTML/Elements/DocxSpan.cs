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
		
		internal override bool CanConvert(HtmlNode node)
		{
			return string.Compare(node.Tag, "span", true) == 0;
		}
		
		internal override void Process(HtmlNode node, OpenXmlElement parent)
		{
			if (node != null && parent != null)
			{
				Run run = null;
				
				foreach (HtmlNode child in node.Children)
				{
					if (child.IsText)
					{
						if (run == null)
						{
							run = AppendRun(parent);
						}
						
						run.AppendChild(new Text(node.InnerHtml));
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

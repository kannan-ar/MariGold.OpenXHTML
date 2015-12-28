namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxBold : DocxElement
	{
		public DocxBold(IOpenXmlContext context)
			:base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "b", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent)
		{
			if (node == null)
			{
				return;
			}
			
			foreach (IHtmlNode child in node.Children)
			{
				if (child.IsText)
				{
					Run run = CreateRun(child);
					
					if (run.RunProperties == null)
					{
						run.RunProperties = new RunProperties();
					}
					
					DocxFont.ApplyBold(run.RunProperties);
					
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

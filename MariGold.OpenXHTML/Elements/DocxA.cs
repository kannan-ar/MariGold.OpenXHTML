namespace MariGold.OpenXHTML
{
	using System;
	using System.Linq;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxA : DocxElement
	{
		private const string href = "href";
		
		internal DocxA(IOpenXmlContext context)
			: base(context)
		{
			
		}
		
		internal override bool CanConvert(HtmlNode node)
		{
			return string.Compare(node.Tag, "a", true) == 0;
		}
		
		internal override void Process(HtmlNode node, OpenXmlElement parent)
		{
			if (node == null)
			{
				return;
			}
			
			string link = ExtractAttributeValue(href, node);
			
			if (!string.IsNullOrEmpty(link))
			{
				Uri uri = new Uri(link);
				
				var relationship = context.MainDocumentPart.AddHyperlinkRelationship(uri, uri.IsAbsoluteUri);
				
				var hyperLink = new Hyperlink() { History = true, Id = relationship.Id };
				
				Run run = new Run();
				run.RunProperties = new RunProperties((new RunStyle() { Val = "Hyperlink" }));
				
				foreach (HtmlNode child in node.Children)
				{
					if (child.IsText && !string.IsNullOrEmpty(child.InnerHtml))
					{
						run.AppendChild(new Text(child.InnerHtml));
					}
				}
				
				hyperLink.Append(run);
				
				if (parent is Paragraph)
				{
					parent.Append(hyperLink);
				}
				else
				{
					if (context.LastParagraph == null)
					{
						context.LastParagraph = parent.AppendChild(new Paragraph());
					}
					
					context.LastParagraph.Append(hyperLink);
				}
			}
		}
	}
}

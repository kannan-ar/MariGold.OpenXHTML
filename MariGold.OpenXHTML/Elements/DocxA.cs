namespace MariGold.OpenXHTML
{
	using System;
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
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "a", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node == null)
			{
				return;
			}
			
			DocxNode docxNode = new DocxNode(node);
			
			string link = docxNode.ExtractAttributeValue(href);
			
			if (Uri.IsWellFormedUriString(link, UriKind.Relative) && !string.IsNullOrEmpty(context.BaseURL))
			{
				link = string.Concat(context.BaseURL, link);
			}
			
			if (Uri.IsWellFormedUriString(link, UriKind.Absolute))
			{
				Uri uri = new Uri(link);
				
				var relationship = context.MainDocumentPart.AddHyperlinkRelationship(uri, uri.IsAbsoluteUri);
				
				var hyperLink = new Hyperlink() { History = true, Id = relationship.Id };
				
				Run run = new Run();
				RunCreated(node, run);
				
				if (run.RunProperties == null)
				{
					run.RunProperties = new RunProperties((new RunStyle() { Val = "Hyperlink" }));
				}
				else
				{
					run.RunProperties.Append(new RunStyle() { Val = "Hyperlink" });
				}
				
				foreach (IHtmlNode child in node.Children)
				{
					if (child.IsText && !IsEmptyText(child.InnerHtml))
					{
						run.AppendChild(new Text() {
							Text = ClearHtml(child.InnerHtml),
							Space = SpaceProcessingModeValues.Preserve
						});
					}
					else
					{
						ProcessTextElement(child, hyperLink);
					}
				}
				
				hyperLink.Append(run);
				
				if (paragraph == null)
				{
					paragraph = parent.AppendChild(new Paragraph());
					ParagraphCreated(node, paragraph);
				}
				
				paragraph.Append(hyperLink);
			}
		}
	}
}

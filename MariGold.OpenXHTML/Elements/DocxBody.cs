namespace MariGold.OpenXHTML
{
	using System;
	using System.Linq;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxBody : DocxElement
	{
		private OpenXmlElement body;
		
		private void ProcessBody(IHtmlNode node, ref Paragraph paragraph)
		{
			while (node != null)
			{
				if (node.IsText && !IsEmptyText(node.InnerHtml))
				{
					if (paragraph == null)
					{
						paragraph = body.AppendChild(new Paragraph());
						ParagraphCreated(node, paragraph);
					}
					
					Run run = paragraph.AppendChild(new Run(new Text(ClearHtml(node.InnerHtml))));
					RunCreated(node, run);
				}
				else
				{
					ProcessChild(node, body, ref paragraph);
				}
				
				node = node.Next;
			}
		}
		
		public DocxBody(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "body", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			body = context.Document.AppendChild(new Body());
			
			//If the node is body tag, find the first children to process
			if (CanConvert(node))
			{
				if (!node.HasChildren)
				{
					//Nothing to process. Just return from here.
					return;
				}
				
				node = node.Children.ElementAt(0);
			}
			
			ProcessBody(node, ref paragraph);
		}
	}
}

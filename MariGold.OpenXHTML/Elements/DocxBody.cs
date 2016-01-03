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
		
		private void ProcessBody(IHtmlNode node)
		{
			while (node != null)
			{
				if (node.IsText)
				{
					AppendToParagraphWithRun(node, body, new Text(node.InnerHtml));
				}
				else
				{
					ProcessChild(node, body);
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
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent)
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
			
			ProcessBody(node);
		}
	}
}

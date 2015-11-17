namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxBody : WordElement
	{
		private void ProcessBody(HtmlNode node)
		{
			OpenXmlElement body = context.Document.AppendChild(new Body());
			
			Run run = null;
			
			while (node != null)
			{
				if (node.Tag == "#text")
				{
					if (paragraph == null && run == null)
					{
						run = body.AppendChild(new Run());
					}
					
				}
				else
				{
					//Reset the run to finilize the text area and restart after appending the current node.
					run = null;
					
					ProcessChild(node, body);
				}
				
				node = node.Next;
			}
		}
		
		public DocxBody(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool IsBlockLine
		{
			get
			{
				return true;
			}
		}
		
		internal override bool CanConvert(HtmlNode node)
		{
			return string.Compare(node.Tag, "body", true) == 0;
		}
		
		internal override void Process(HtmlNode node, OpenXmlElement parent)
		{
			//If the node is body tag, find the first children to process
			if (CanConvert(node))
			{
				if (!node.HasChildren)
				{
					//Nothing to process. Just return from here.
					return;
				}
				
				node = node.Children[0];
			}
			
			ProcessBody(node);
		}
	}
}

namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	
	internal abstract class WordElement
	{
		protected readonly IOpenXmlContext context;
		
		protected void ProcessChild(HtmlNode node, OpenXmlElement parent)
		{
			WordElement element = context.Convert(node);
					
			if (element != null)
			{
				element.Process(node, parent);
			}
		}
		
		internal WordElement(IOpenXmlContext context)
		{
			if (context == null)
			{
				throw new ArgumentNullException("context");
			}
			
			this.context = context;
		}
		
		internal abstract bool IsBlockLine{ get; }
		
		internal abstract bool CanConvert(HtmlNode node);
		internal abstract void Process(HtmlNode node, OpenXmlElement parent);
	}
}

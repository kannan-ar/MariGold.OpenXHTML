namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	
	internal abstract class WordElement
	{
		private readonly IWordContext context;
		
		
		internal WordElement(IWordContext context)
		{
			if (context == null)
			{
				throw new ArgumentNullException("context");
			}
			
			this.context = context;
		}
		
		internal abstract bool CanConvert(HtmlNode node);
		internal abstract void Process(HtmlNode node);
	}
}

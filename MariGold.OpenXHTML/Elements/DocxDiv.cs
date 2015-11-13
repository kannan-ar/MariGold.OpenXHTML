namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	
	internal sealed class DocxDiv : WordElement
	{
		internal DocxDiv(IWordContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(HtmlNode node)
		{
			throw new NotImplementedException();
		}
		
		internal override void Process(HtmlNode node)
		{
			throw new NotImplementedException();
		}
	}
}

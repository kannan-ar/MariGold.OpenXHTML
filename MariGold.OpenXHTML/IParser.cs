namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	
	/// <summary>
	/// 
	/// </summary>
	public interface IParser
	{
		IHtmlNode FindBodyOrFirstElement();
	}
}

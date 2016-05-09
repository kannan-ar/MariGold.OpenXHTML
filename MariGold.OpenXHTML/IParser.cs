namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	
	/// <summary>
	/// 
	/// </summary>
	public interface IParser
	{
        string BaseURL { get; set; }
        string UriSchema { get; set; }
		IHtmlNode FindBodyOrFirstElement();
	}
}

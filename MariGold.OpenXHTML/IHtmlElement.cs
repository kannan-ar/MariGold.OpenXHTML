namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	/// <summary>
	/// 
	/// </summary>
	public interface IHtmlElement
	{
		string Tag{ get; }
		string InnerHtml{ get; }
		string Html{ get; }
		IDictionary<string,string> Attributes{ get; }
		IDictionary<string,string> Styles{ get; }
	}
}

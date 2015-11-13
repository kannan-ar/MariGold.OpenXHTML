namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	
	/// <summary>
	/// 
	/// </summary>
	public class HtmlParser : IParser
	{
		private readonly string html;
		
		private HtmlNode FindBody(HtmlNode node)
		{
			if (string.Compare(node.Tag, "body", true) == 0)
			{
				return node;
			}
			
			foreach (HtmlNode child in node.Children)
			{
				HtmlNode body = FindBody(child);
				
				if (body != null)
				{
					return body;
				}
			}
			
			return null;
		}
		
		public HtmlParser(string html)
		{
			this.html = html;
		}
		
		public HtmlNode FindBodyOrFirstElement()
		{
			MariGold.HtmlParser.HtmlParser parser = new HtmlTextParser(html);
			
			parser.Parse();
			parser.ParseCSS();
			
			HtmlNode node = parser.Current;
			HtmlNode body = null;
			
			while (node != null)
			{
				body = FindBody(node);
				
				if (body != null)
				{
					break;
				}
				else if (node.Next != null)
				{
					node = node.Next;
				}
			}
			
			return body ?? parser.Current;
		}
	}
}

namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using System.Collections.Generic;
	
	internal sealed class DocxNode
	{
		private readonly IHtmlNode node;
		
		internal DocxNode(IHtmlNode node)
		{
			if (node == null)
			{
				throw new ArgumentNullException("node");
			}
			
			this.node = node;
		}
		
		internal string ExtractAttributeValue(string attributeName)
		{
			foreach (KeyValuePair<string,string> attribute in node.Attributes)
			{
				if (string.Compare(attributeName, attribute.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					return attribute.Value;
				}
			}
			
			return string.Empty;
		}
		
		internal string ExtractStyleValue(string styleName)
		{
			foreach (KeyValuePair<string,string> style in node.Styles)
			{
				if (string.Compare(styleName, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					return style.Value;
				}
			}
			
			return string.Empty;
		}
		
		internal void SetStyleValue(string styleName, string value)
		{
			string key = string.Empty;
			
			foreach (KeyValuePair<string,string> style in node.Styles)
			{
				if (string.Compare(styleName, style.Key, StringComparison.InvariantCultureIgnoreCase) == 0)
				{
					key = style.Key;
				}
			}
			
			if (!string.IsNullOrEmpty(key))
			{
				node.Styles[key] = value;
			}
			else
			{
				node.Styles[styleName] = value;
			}
		}
	}
}

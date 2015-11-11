namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	
	/// <summary>
	/// 
	/// </summary>
	public class HtmlElement : IHtmlElement
	{
		private readonly string tag;
		private readonly string innerHtml;
		private readonly string html;
		private readonly IDictionary<string,string> attributes;
		private readonly IDictionary<string,string> styles;
		
		public HtmlElement(
			string tag, 
			string innerHtml,
			string html,
			IDictionary<string,string> attributes,
			IDictionary<string,string> styles)
		{
			this.tag = tag;
			this.innerHtml = innerHtml;
			this.html = html;
			this.attributes = attributes;
			this.styles = styles;
		}
		
		public string Tag
		{
			get
			{
				return tag;
			}
		}
		
		public string InnerHtml
		{
			get
			{
				return innerHtml;
			}
		}
		
		public string Html
		{
			get
			{
				return html;
			}
		}
		
		public IDictionary<string,string> Attributes
		{
			get
			{
				return attributes;
			}
		}
		
		public IDictionary<string,string> Styles
		{
			get
			{
				return styles;
			}
		}
	}
}

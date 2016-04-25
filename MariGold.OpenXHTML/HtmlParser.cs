namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	
	/// <summary>
	/// 
	/// </summary>
	public sealed class HtmlParser : IParser
	{
		private readonly string html;

        private string uriSchema;

		private IHtmlNode FindBody(IHtmlNode node)
		{
			if (string.Compare(node.Tag, "body", StringComparison.InvariantCultureIgnoreCase) == 0)
			{
				return node;
			}
			
			foreach (IHtmlNode child in node.Children)
			{
				IHtmlNode body = FindBody(child);
				
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
		
        public string UriSchema
        {
            get
            {
                return uriSchema;
            }

            set
            {
                uriSchema = value;
            }
        }

		public IHtmlNode FindBodyOrFirstElement()
		{
			MariGold.HtmlParser.HtmlParser parser = new HtmlTextParser(html);

            parser.UriSchema = uriSchema;
			parser.Parse();
			parser.ParseCSS();
			
			IHtmlNode node = parser.Current;
			IHtmlNode body = null;
			
			while (node != null)
			{
				body = FindBody(node);
				
				if (body != null || node.Next == null)
				{
					break;
				}
				
				node = node.Next;
			}
			
			return body ?? parser.Current;
		}
	}
}

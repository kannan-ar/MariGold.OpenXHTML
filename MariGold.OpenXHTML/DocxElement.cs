namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal abstract class DocxElement
	{
		protected readonly IOpenXmlContext context;
		
		protected void RunCreated(IHtmlNode node, Run run)
		{
			DocxRunStyle style = new DocxRunStyle();
			style.Process(run, node);
		}
		
		protected void ParagraphCreated(IHtmlNode node, Paragraph para)
		{
			DocxParagraphStyle style = new DocxParagraphStyle();
			style.Process(para, node);
		}
		
		protected void ProcessChild(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			DocxElement element = context.Convert(node);
					
			if (element != null)
			{
				element.Process(node, parent, ref paragraph);
			}
		}
		
		internal DocxElement(IOpenXmlContext context)
		{
			if (context == null)
			{
				throw new ArgumentNullException("context");
			}
			
			this.context = context;
		}
		
		internal string ClearHtml(string html)
		{
			if (string.IsNullOrEmpty(html))
			{
				return string.Empty;
			}
			
			html = html.Replace("&nbsp;", " ");
			html = html.Replace("&amp;", "&");
			return html.Replace(Environment.NewLine, string.Empty);
		}
		
		internal bool IsEmptyText(string html)
		{
			if (string.IsNullOrEmpty(html))
			{
				return true;
			}
			
			html = html.Replace(Environment.NewLine, string.Empty);
			
			if (string.IsNullOrEmpty(html))
			{
				return true;
			}
			
			return false;
		}
		
		internal abstract bool CanConvert(IHtmlNode node);
		
		internal abstract void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph);
	}
}

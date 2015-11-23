namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal abstract class DocxElement
	{
		protected readonly IOpenXmlContext context;
		
		protected void ProcessChild(HtmlNode node, OpenXmlElement parent)
		{
			DocxElement element = context.Convert(node);
					
			if (element != null)
			{
				element.Process(node, parent);
			}
		}
		
		protected string ExtractAttributeValue(string attributeName, HtmlNode node)
		{
			if (node == null)
			{
				return string.Empty;
			}
			
			foreach (KeyValuePair<string,string> attribute in node.Attributes)
			{
				if (string.Compare(attributeName, attribute.Key) == 0)
				{
					return attribute.Value;
				}
			}
			
			return string.Empty;
		}
		
		protected void AppendToParagraph(OpenXmlElement parent, OpenXmlElement element)
		{
			if (parent is Paragraph)
			{
				parent.Append(element);
			}
			else
			{
				if (context.LastParagraph == null)
				{
					context.LastParagraph = parent.AppendChild(new Paragraph());
				}
					
				context.LastParagraph.Append(element);
			}
		}
		
		protected void AppendToParagraphWithRun(OpenXmlElement parent, OpenXmlElement element)
		{
			if (parent is Paragraph)
			{
				parent.Append(new Run(element));
			}
			else
			{
				if (context.LastParagraph == null)
				{
					context.LastParagraph = parent.AppendChild(new Paragraph());
				}
					
				context.LastParagraph.Append(new Run(element));
			}
		}
		
		protected Run AppendRun(OpenXmlElement parent)
		{
			Run run = null;
			
			if (parent is Paragraph)
			{
				run = parent.AppendChild(new Run());
			}
			else
			{
				if (context.LastParagraph == null)
				{
					context.LastParagraph = parent.AppendChild(new Paragraph());
				}
								
				run = context.LastParagraph.AppendChild(new Run());
			}
			
			return run;
		}
		
		internal DocxElement(IOpenXmlContext context)
		{
			if (context == null)
			{
				throw new ArgumentNullException("context");
			}
			
			this.context = context;
		}
		
		internal abstract bool CanConvert(HtmlNode node);
		internal abstract void Process(HtmlNode node, OpenXmlElement parent);
	}
}

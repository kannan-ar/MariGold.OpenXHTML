namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal abstract class DocxElement
	{
		private Paragraph paragraph;
		private DocxElement parent;
		
		protected readonly IOpenXmlContext context;
		
		private void RunCreated(HtmlNode node, Run run)
		{
			DocxRunStyle style = new DocxRunStyle();
			style.Process(run, node.Styles);
		}
		
		private void ParagraphCreated(HtmlNode node, Paragraph para)
		{
			DocxParagraphStyle style = new DocxParagraphStyle();
			style.Process(para, node.Styles);
		}
		
		internal Paragraph Current
		{
			get
			{
				return paragraph;
			}
			
			set
			{
				paragraph = value;
			}
		}
		
		internal DocxElement Parent
		{
			get
			{
				return parent;
			}
			
			set
			{
				parent = value;
			}
		}
		
		protected void ProcessChild(HtmlNode node, OpenXmlElement parent)
		{
			DocxElement element = context.Convert(node);
					
			if (element != null)
			{
				element.Parent = this;
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
		
		protected void AppendToParagraph(HtmlNode node, OpenXmlElement parent, OpenXmlElement element)
		{
			if (parent is Paragraph)
			{
				parent.Append(element);
			}
			else
			{
				if (Parent.Current == null)
				{
					Parent.Current = parent.AppendChild(new Paragraph());
					ParagraphCreated(node, Parent.Current);
				}
					
				Parent.Current.Append(element);
			}
		}
		
		protected void AppendToParagraphWithRun(HtmlNode node, OpenXmlElement parent, OpenXmlElement element)
		{
			if (parent is Paragraph)
			{
				Run run = new Run(element);
				parent.Append(run);
				RunCreated(node, run);
			}
			else
			{
				if (Parent.Current == null)
				{
					Parent.Current = parent.AppendChild(new Paragraph());
					ParagraphCreated(node, Parent.Current);
				}
				
				Run run = new Run(element);
				Parent.Current.Append(run);
				RunCreated(node, run);
			}
		}
		
		protected Run AppendRun(HtmlNode node, OpenXmlElement parent)
		{
			Run run = null;
			
			if (parent is Paragraph)
			{
				run = parent.AppendChild(new Run());
				RunCreated(node, run);
			}
			else
			{
				if (Parent.Current == null)
				{
					Parent.Current = parent.AppendChild(new Paragraph());
					ParagraphCreated(node, Parent.Current);
				}
								
				run = Parent.Current.AppendChild(new Run());
				RunCreated(node, run);
			}
			
			return run;
		}
		
		protected Run CreateRun(HtmlNode node)
		{
			Run run = new Run();
			RunCreated(node, run);
			return run;
		}
		
		protected Paragraph CreateParagraph(HtmlNode node)
		{
			Paragraph para = new Paragraph();
			ParagraphCreated(node, para);
			return para;
		}
		
		protected Paragraph CreateParagraph(HtmlNode node, OpenXmlElement parent)
		{
			Paragraph para = parent.AppendChild(new Paragraph());
			ParagraphCreated(node, para);
			return para;
		}
		
		protected Run CreateRun(HtmlNode node, OpenXmlElement parent)
		{
			Run run = parent.AppendChild(new Run());
			RunCreated(node, run);
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

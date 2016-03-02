namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.Text.RegularExpressions;
	
	internal sealed class DocxHeader : DocxElement
	{
		private Regex isValid;
		
		private int GetHeaderNumber(IHtmlNode node)
		{
			int value = -1;
			Regex regex = new Regex("[1-6]{1}$");
			
			Match match = regex.Match(node.Tag);
			
			if (match != null)
			{
				Int32.TryParse(match.Value, out value);
			}
			
			return value;
		}
		
		private int CalculateFontSize(int headerSize)
		{
			int fontSize = -1;
			
			switch (headerSize)
			{
				case 1:
					fontSize = 32;
					break;
					
				case 2:
					fontSize = 24;
					break;
					
				case 3:
					fontSize = 19;
					break;
					
				case 4:
					fontSize = 22;
					break;
					
				case 5:
					fontSize = 13;
					break;
					
				case 6:
					fontSize = 11;
					break;
			}
			
			return fontSize;
		}
		
		private void ApplyStyle(IHtmlNode node, Run run)
		{
			int fontSize = CalculateFontSize(GetHeaderNumber(node));
			
			if (fontSize == -1)
			{
				return;
			}
			
			if (run.RunProperties == null)
			{
				run.RunProperties = new RunProperties();
			}
			
			DocxFont.ApplyFont(fontSize, true, run.RunProperties);
		}
		
		internal DocxHeader(IOpenXmlContext context)
			: base(context)
		{
			isValid = new Regex(@"^[hH][1-6]{1}$");
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return isValid.IsMatch(node.Tag);
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node != null && parent != null)
			{
				paragraph = null;
				Paragraph headerParagraph = null;
				
				foreach (IHtmlNode child in node.Children)
				{
					if (child.IsText)
					{
						if (headerParagraph == null)
						{
							headerParagraph = parent.AppendChild(new Paragraph());
							ParagraphCreated(node, headerParagraph);
						}
						
						Run run = headerParagraph.AppendChild(new Run());
						RunCreated(child, run);
						ApplyStyle(node, run);
						
						run.AppendChild(new Text(child.InnerHtml));
					}
					else
					{
						ProcessChild(child, parent, ref headerParagraph);
					}
				}
			}
		}
	}
}

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
		
		private int GetHeaderNumber(HtmlNode node)
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
		
		private void ApplyStyle(HtmlNode node, Run run)
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
		
		internal override bool CanConvert(HtmlNode node)
		{
			return isValid.IsMatch(node.Tag);
		}
		
		internal override void Process(HtmlNode node, OpenXmlElement parent)
		{
			if (node != null && parent != null) 
			{
				Parent.Current = null;
				OpenXmlElement paragraph = CreateParagraph(node, parent);
			
				foreach (HtmlNode child in node.Children) 
				{
					if (child.IsText) 
					{
						Run run = AppendRun(node, paragraph);
						
						ApplyStyle(node, run);
						
						run.AppendChild(new Text(node.InnerHtml));
					}
					else 
					{
						ProcessChild(child, parent);
					}
				}
			}
		}
	}
}

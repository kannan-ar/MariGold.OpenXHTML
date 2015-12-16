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
		
		internal DocxHeader(IOpenXmlContext context)
			:base(context)
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

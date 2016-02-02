namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxBold : DocxElement
	{
		public DocxBold(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "b", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node == null)
			{
				return;
			}
			
			foreach (IHtmlNode child in node.Children)
			{
				if (child.IsText)
				{
					if (paragraph == null)
					{
						paragraph = parent.AppendChild(new Paragraph());
						ParagraphCreated(child, paragraph);
					}
					
					Run run = paragraph.AppendChild(new Run());
					RunCreated(child, run);
					
					/*
					Run run = CreateRun(child);
					
					
					*/
					
					//Need to analyze the child style properties. If there is a bold-weight:normal property, 
					//apply bold should not happen
					if (run.RunProperties == null)
					{
						run.RunProperties = new RunProperties();
					}
					
					DocxFont.ApplyBold(run.RunProperties);
					
					run.AppendChild(new Text(child.InnerHtml));
					                
					//AppendToParagraph(node, parent, run);
					paragraph.Append(run);
				}
				else
				{
					ProcessChild(child, parent, ref paragraph);
				}
			}
		}
	}
}

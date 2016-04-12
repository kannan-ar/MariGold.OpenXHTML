namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxBold : DocxElement, ITextElement
	{
		private void SetStyle(IHtmlNode node)
		{
			DocxNode docxNode = new DocxNode(node);
			
			string value = docxNode.ExtractStyleValue(DocxFont.fontWeight);
			
			if (string.IsNullOrEmpty(value))
			{
				docxNode.SetStyleValue(DocxFont.fontWeight, DocxFont.bold);
			}
		}
		
		private void ProcessRun(Run run, IHtmlNode child)
		{
			RunCreated(child, run);
					
			//Need to analyze the child style properties. If there is a bold-weight:normal property, 
			//apply bold should not happen
			if (run.RunProperties == null)
			{
				run.RunProperties = new RunProperties();
			}
					
			DocxFont.ApplyBold(run.RunProperties);
					
			run.AppendChild(new Text() {
				Text = ClearHtml(child.InnerHtml),
				Space = SpaceProcessingModeValues.Preserve						
			});
		}
		
		public DocxBold(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "b", StringComparison.InvariantCultureIgnoreCase) == 0 ||
			string.Compare(node.Tag, "strong", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node == null)
			{
				return;
			}
			
			foreach (IHtmlNode child in node.Children)
			{
				if (child.IsText && !IsEmptyText(child.InnerHtml))
				{
					if (paragraph == null)
					{
						paragraph = parent.AppendChild(new Paragraph());
						IHtmlNode parentNode = node.Parent ?? node;
						ParagraphCreated(parentNode, paragraph);
					}
					
					Run run = paragraph.AppendChild(new Run());
					ProcessRun(run, child);
				}
				else
				{
					SetStyle(child);
					ProcessChild(child, parent, ref paragraph);
				}
			}
		}
		
		bool ITextElement.CanConvert(IHtmlNode node)
		{
			return CanConvert(node);
		}
		
		void ITextElement.Process(IHtmlNode node, OpenXmlElement parent)
		{
			foreach (IHtmlNode child in node.Children)
			{
				if (child.IsText && !IsEmptyText(child.InnerHtml))
				{
					Run run = parent.AppendChild(new Run());
					ProcessRun(run, child);
				}
				else
				{
					SetStyle(child);
					ProcessTextElement(child, parent);
				}
			}
		}
	}
}

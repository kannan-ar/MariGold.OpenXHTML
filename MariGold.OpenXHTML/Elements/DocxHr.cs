namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxHr : DocxElement
	{
		internal DocxHr(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(DocxNode node)
		{
			return string.Compare(node.Tag, "hr", StringComparison.InvariantCultureIgnoreCase) == 0;
		}

        internal override void Process(DocxNode node, ref Paragraph paragraph)
		{
            if (node.IsNull() || node.Parent == null || IsHidden(node))
			{
				return;
			}
			
			paragraph = null;

            Paragraph hrParagraph = node.Parent.AppendChild(new Paragraph());
            OnParagraphCreated(node, hrParagraph);
				
			if (hrParagraph.ParagraphProperties == null)
			{
				hrParagraph.ParagraphProperties = new ParagraphProperties();
			}
				
			ParagraphBorders paragraphBorders = new ParagraphBorders();
			DocxBorder.ApplyDefaultBorder<TopBorder>(paragraphBorders);
			hrParagraph.ParagraphProperties.Append(paragraphBorders);
				
			Run run = hrParagraph.AppendChild(new Run(new Text()));
            RunCreated(node, run);
		}
	}
}

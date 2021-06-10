namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections.Generic;
    using DocumentFormat.OpenXml.Vml;
    using DocumentFormat.OpenXml.Vml.Office;
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

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
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

			var rectangle = new Rectangle();
			rectangle.Style = "width:0;height:1.5pt";
			rectangle.Horizontal = true;
			rectangle.HorizontalStandard = true;
			rectangle.FillColor = "#a0a0a0";
			rectangle.Stroked = false;
			rectangle.HorizontalAlignment = HorizontalRuleAlignmentValues.Center;
			var picture = new Picture(rectangle);

			Run run = hrParagraph.AppendChild(new Run(picture));
            RunCreated(node, run);
		}
	}
}

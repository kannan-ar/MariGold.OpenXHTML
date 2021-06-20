namespace MariGold.OpenXHTML
{
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxParagraphStyle
	{
		private void ProcessBorder(DocxNode node, ParagraphProperties properties)
		{
			ParagraphBorders paragraphBorders = new ParagraphBorders();
			
			DocxBorder.ApplyBorders(paragraphBorders,
				node.ExtractStyleValue(DocxBorder.borderName),
				node.ExtractStyleValue(DocxBorder.leftBorderName),
				node.ExtractStyleValue(DocxBorder.topBorderName),
				node.ExtractStyleValue(DocxBorder.rightBorderName),
				node.ExtractStyleValue(DocxBorder.bottomBorderName),
				false);
			
			if (paragraphBorders.HasChildren)
			{
				properties.Append(paragraphBorders);
			}
		}

		static internal void SetIndentation(Paragraph element, int indent)
		{
			if (element.ParagraphProperties == null)
			{
				element.ParagraphProperties = new ParagraphProperties();
			}

			element.ParagraphProperties.Append(new Indentation() { Left = indent.ToString() });
		}

		private void CheckFonts(DocxNode node, ParagraphMarkRunProperties properties)
		{
			string fontFamily = node.ExtractStyleValue(DocxFontStyle.fontFamily);
			string fontWeight = node.ExtractStyleValue(DocxFontStyle.fontWeight);
			string fontStyle = node.ExtractStyleValue(DocxFontStyle.fontStyle);

			if (!string.IsNullOrEmpty(fontFamily))
			{
				DocxFontStyle.ApplyFontFamily(fontFamily, properties);
			}

			if (!string.IsNullOrEmpty(fontWeight))
			{
				DocxFontStyle.ApplyFontWeight(fontWeight, properties);
			}

			if (!string.IsNullOrEmpty(fontStyle))
			{
				DocxFontStyle.ApplyFontStyle(fontStyle, properties);
			}
		}

		private void CheckFontStyle(DocxNode node, ParagraphMarkRunProperties properties)
		{
			string fontSize = node.ExtractStyleValue(DocxFontStyle.fontSize);
			string textDecoration = node.ExtractStyleValue(DocxFontStyle.textDecoration);

			if (string.IsNullOrEmpty(textDecoration))
			{
				textDecoration = node.ExtractStyleValue(DocxFontStyle.textDecorationLine);
			}

			if (!string.IsNullOrEmpty(fontSize))
			{
				DocxFontStyle.ApplyFontSize(fontSize, properties);
			}

			if (!string.IsNullOrEmpty(textDecoration))
			{
				DocxFontStyle.ApplyTextDecoration(textDecoration, properties);
			}
		}

		private void ProcessDirection(DocxNode node, ParagraphProperties properties)
        {
			string styleValue = node.ExtractStyleValue(DocxDirection.direction);

			if (!string.IsNullOrEmpty(styleValue))
			{
				DocxDirection.ApplyBidi(styleValue, properties);
			}
		}

		internal void Process(Paragraph element, DocxNode node)
		{
			ParagraphProperties properties = element.ParagraphProperties;
			
			if (properties == null)
			{
				properties = new ParagraphProperties();
			}
			
			//Order of assigning styles to paragraph property is important. The order should not change.
            ProcessBorder(node, properties);

            string backgroundColor = node.ExtractStyleValue(DocxColor.backGroundColor);
            string backGround = DocxColor.ExtractBackGround(node.ExtractStyleValue(DocxColor.backGround));

			if (!string.IsNullOrEmpty(backgroundColor))
			{
				DocxColor.ApplyBackGroundColor(backgroundColor, properties);
			}
			else if(!string.IsNullOrEmpty(backGround))
            {
                DocxColor.ApplyBackGroundColor(backGround, properties);
            }

            DocxMargin margin = new DocxMargin(node);
			margin.ProcessParagraphMargin(properties);

            string textAlign = node.ExtractStyleValue(DocxAlignment.textAlign);
			if (!string.IsNullOrEmpty(textAlign))
			{
				DocxAlignment.ApplyTextAlign(textAlign, properties);
			}

			ProcessDirection(node, properties);

			#region Set Run Properties for the Paragraph Mark
			var runProperties = properties.ParagraphMarkRunProperties;

			if (runProperties == null)
			{
				runProperties = new ParagraphMarkRunProperties();
			}

			//Order of assigning styles to run property is important. The order should not change.
			CheckFonts(node, runProperties);

			string color = node.ExtractStyleValue(DocxColor.color);

			if (!string.IsNullOrEmpty(color))
			{
				DocxColor.ApplyColor(color, runProperties);
			}

			CheckFontStyle(node, runProperties);

			if (properties.ParagraphMarkRunProperties == null && runProperties.HasChildren)
			{
				properties.ParagraphMarkRunProperties = runProperties;
			}
			#endregion


			if (element.ParagraphProperties == null && properties.HasChildren)
			{
				element.ParagraphProperties = properties;
			}
		}
	}
}

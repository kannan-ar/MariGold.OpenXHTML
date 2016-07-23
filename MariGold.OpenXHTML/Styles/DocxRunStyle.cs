namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal sealed class DocxRunStyle
	{
		private void CheckFonts(DocxNode node, RunProperties properties)
		{
            string fontFamily = node.ExtractStyleValue(DocxFont.fontFamily);
            string fontWeight = node.ExtractStyleValue(DocxFont.fontWeight);
            string fontStyle = node.ExtractStyleValue(DocxFont.fontStyle);
			
			if (!string.IsNullOrEmpty(fontFamily))
			{
				DocxFont.ApplyFontFamily(fontFamily, properties);
			}
			
			if (!string.IsNullOrEmpty(fontWeight))
			{
				DocxFont.ApplyFontWeight(fontWeight, properties);
			}
				
			if (!string.IsNullOrEmpty(fontStyle))
			{
				DocxFont.ApplyFontStyle(fontStyle, properties);
			}
		}

        private void CheckFontStyle(DocxNode node, RunProperties properties)
		{
            string fontSize = node.ExtractStyleValue(DocxFont.fontSize);
            string textDecoration = node.ExtractStyleValue(DocxFont.textDecoration);
			
			if (!string.IsNullOrEmpty(fontSize))
			{
				DocxFont.ApplyFontSize(fontSize, properties);
			}
			
			if (!string.IsNullOrEmpty(textDecoration))
			{
				DocxFont.ApplyTextDecoration(textDecoration, properties);
			}
		}

        private void ProcessBackGround(DocxNode node, RunProperties properties)
        {
            string backgroundColor = node.ExtractStyleValue(DocxColor.backGroundColor);
            string backGround = DocxColor.ExtractBackGround(node.ExtractStyleValue(DocxColor.backGround));

            if (!string.IsNullOrEmpty(backgroundColor))
            {
                DocxColor.ApplyBackGroundColor(backgroundColor, properties);
            }
            else if (!string.IsNullOrEmpty(backGround))
            {
                DocxColor.ApplyBackGroundColor(backGround, properties);
            }
        }

        internal void Process(Run element, DocxNode node)
		{
			RunProperties properties = element.RunProperties;
			
			if (properties == null)
			{
				properties = new RunProperties();
			}
			
			//Order of assigning styles to run property is important. The order should not change.
            CheckFonts(node, properties);

            string color = node.ExtractStyleValue(DocxColor.color);
			
			if (!string.IsNullOrEmpty(color))
			{
				DocxColor.ApplyColor(color, properties);
			}

            CheckFontStyle(node, properties);

            ProcessBackGround(node, properties);
			
			if (element.RunProperties == null && properties.HasChildren)
			{
				element.RunProperties = properties;
			}
		}
	}
}

namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxOL : DocxElement
	{
		private const string elementName = "ol";
		private const string liName = "li";
		
		private void InitNumberDefinitions(NumberFormatValues numberFormat)
		{
			if (!context.HasNumberingDefinition(numberFormat))
			{
				//Enum values starting from zero. We need non zero values here
				Int32 numberId = ((Int32)numberFormat) + 1;
				
				AbstractNum abstractNum = new AbstractNum() { AbstractNumberId = numberId };

				Level level = new Level() { LevelIndex = 0 };
				StartNumberingValue startNumberingValue = new StartNumberingValue() { Val = 1 };
				NumberingFormat numberingFormat = new NumberingFormat() { Val = numberFormat };
				LevelText levelText = new LevelText() { Val = "%1." };
				LevelJustification levelJustification = new LevelJustification() { Val = LevelJustificationValues.Left };

				PreviousParagraphProperties previousParagraphProperties = new PreviousParagraphProperties();
				Indentation indentation = new Indentation() { Start = "720", Hanging = "360" };

				previousParagraphProperties.Append(indentation);

				level.Append(startNumberingValue);
				level.Append(numberingFormat);
				level.Append(levelText);
				level.Append(levelJustification);
				level.Append(previousParagraphProperties);
            
				abstractNum.Append(level);
				
				NumberingInstance numberingInstance = new NumberingInstance() { NumberID = numberId };
				AbstractNumId abstractNumId = new AbstractNumId() { Val = numberId };

				numberingInstance.Append(abstractNumId);
				
				context.SaveNumberingDefinition(numberFormat, abstractNum, numberingInstance);
			}
		}
		
		private ParagraphProperties GetListProperties(NumberFormatValues numberFormat)
		{
			ParagraphProperties paragraphProperties = new ParagraphProperties();
			ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };

			NumberingProperties numberingProperties = new NumberingProperties();
			NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = 0 };
			NumberingId numberingId = new NumberingId() { Val = ((Int32)numberFormat) + 1 };

			numberingProperties.Append(numberingLevelReference);
			numberingProperties.Append(numberingId);

			paragraphProperties.Append(paragraphStyleId);
			paragraphProperties.Append(numberingProperties);

			return paragraphProperties;
		}
		
		private void ProcessLi(IHtmlNode li, OpenXmlElement parent, NumberFormatValues numberFormat)
		{
			OpenXmlElement paragraph = CreateParagraph(li, parent);
			paragraph.Append(GetListProperties(numberFormat));
			
			foreach (IHtmlNode child in li.Children)
			{
				if (child.IsText)
				{
					AppendRun(li, paragraph).AppendChild(new Text(child.InnerHtml));
				}
				else
				{
					ProcessChild(child, paragraph);
				}
			}
		}
		
		private NumberFormatValues GetNumberFormat(IHtmlNode node)
		{
			NumberFormatValues numberFormat = NumberFormatValues.Decimal;
			
			string type;
			
			if (node.Attributes.TryGetValue("type", out type))
			{
				switch (type)
				{
					case "1":
						numberFormat = NumberFormatValues.Decimal;
						break;
						
					case "a":
						numberFormat = NumberFormatValues.LowerLetter;
						break;
						
					case "A":
						numberFormat = NumberFormatValues.UpperLetter;
						break;
						
					case "i":
						numberFormat = NumberFormatValues.LowerRoman;
						break;
						
					case "I":
						numberFormat = NumberFormatValues.UpperRoman;
						break;
				}
			}
			
			return numberFormat;
		}
		
		internal DocxOL(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, elementName, StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent)
		{
			if (node == null || !CanConvert(node))
			{
				return;
			}
			
			if (node.HasChildren)
			{
				NumberFormatValues numberFormat = GetNumberFormat(node);
				
				InitNumberDefinitions(numberFormat);
				
				foreach (IHtmlNode child in node.Children)
				{
					if (string.Compare(child.Tag, liName, StringComparison.InvariantCultureIgnoreCase) == 0)
					{
						ProcessLi(child, parent, numberFormat);
					}
				}
			}
		}
	}
}

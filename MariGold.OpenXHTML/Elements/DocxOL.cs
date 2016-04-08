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
				Indentation indentation = new Indentation() {
					Start = "720",
					Hanging = "360"
				};

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
		
		private void SetListProperties(NumberFormatValues numberFormat, ParagraphProperties paragraphProperties)
		{
			ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };

			NumberingProperties numberingProperties = new NumberingProperties();
			NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = 0 };
			NumberingId numberingId = new NumberingId() { Val = ((Int32)numberFormat) + 1 };

			numberingProperties.Append(numberingLevelReference);
			numberingProperties.Append(numberingId);

			paragraphProperties.Append(paragraphStyleId);
			paragraphProperties.Append(numberingProperties);
		}
		
		private void ProcessLi(IHtmlNode li, OpenXmlElement parent, NumberFormatValues numberFormat)
		{
			Paragraph paragraph = parent.AppendChild(new Paragraph());
			ParagraphCreated(li, paragraph);
			
			if (paragraph.ParagraphProperties == null)
			{
				paragraph.ParagraphProperties = new ParagraphProperties();
			}
			
			SetListProperties(numberFormat, paragraph.ParagraphProperties);
			
			foreach (IHtmlNode child in li.Children)
			{
				if (child.IsText && !IsEmptyText(child.InnerHtml))
				{
					Run run = paragraph.AppendChild(new Run(new Text() {
						Text = ClearHtml(child.InnerHtml),
						Space = SpaceProcessingModeValues.Preserve
					}));
					
					RunCreated(li, run);
				}
				else
				{
					ProcessChild(child, parent, ref paragraph);
				}
			}
		}
		
		private NumberFormatValues GetNumberFormat(IHtmlNode node)
		{
			NumberFormatValues numberFormat = NumberFormatValues.Decimal;
			DocxNode docxNode = new DocxNode(node);
			
			string type = docxNode.ExtractAttributeValue("type");
			
			if (!string.IsNullOrEmpty(type))
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
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node == null || !CanConvert(node))
			{
				return;
			}
			
			paragraph = null;
			
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

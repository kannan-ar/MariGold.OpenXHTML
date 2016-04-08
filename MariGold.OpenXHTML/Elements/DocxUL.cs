namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxUL : DocxElement
	{
		private const string elementName = "ul";
		private const string liName = "li";
		private const NumberFormatValues numberFormat = NumberFormatValues.Bullet;
		
		private void ProcessLi(IHtmlNode li, OpenXmlElement parent)
		{
			Paragraph paragraph = parent.AppendChild(new Paragraph());
			ParagraphCreated(li, paragraph);
			
			if (paragraph.ParagraphProperties == null)
			{
				paragraph.ParagraphProperties = new ParagraphProperties();
			}
			
			SetListProperties(paragraph.ParagraphProperties);
			
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
		
		private void SetListProperties(ParagraphProperties paragraphProperties)
		{
			ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };

			NumberingProperties numberingProperties = new NumberingProperties();
			NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = 0 };
			NumberingId numberingId = new NumberingId() { Val = (Int32)numberFormat };

			numberingProperties.Append(numberingLevelReference);
			numberingProperties.Append(numberingId);

			paragraphProperties.Append(paragraphStyleId);
			paragraphProperties.Append(numberingProperties);
		}
		
		private void InitNumberDefinitions()
		{
			if (!context.HasNumberingDefinition(numberFormat))
			{
				Int32 numberId = (Int32)numberFormat;
				
				AbstractNum abstractNum = new AbstractNum() { AbstractNumberId = numberId };

				Level level = new Level() { LevelIndex = 0 };
				StartNumberingValue startNumberingValue = new StartNumberingValue() { Val = 1 };
				NumberingFormat numberingFormat = new NumberingFormat() { Val = numberFormat };
				LevelText levelText = new LevelText() { Val = "·" };
				LevelJustification levelJustification = new LevelJustification() { Val = LevelJustificationValues.Left };

				PreviousParagraphProperties previousParagraphProperties = new PreviousParagraphProperties();
				Indentation indentation = new Indentation() {
					Start = "720",
					Hanging = "360"
				};

				previousParagraphProperties.Append(indentation);

				NumberingSymbolRunProperties numberingSymbolRunProperties = new NumberingSymbolRunProperties();
				RunFonts runFonts = new RunFonts() {
					Hint = FontTypeHintValues.Default,
					Ascii = "Symbol",
					HighAnsi = "Symbol"
				};

				numberingSymbolRunProperties.Append(runFonts);

				level.Append(startNumberingValue);
				level.Append(numberingFormat);
				level.Append(levelText);
				level.Append(levelJustification);
				level.Append(previousParagraphProperties);
				level.Append(numberingSymbolRunProperties);
            
				abstractNum.Append(level);
				
				NumberingInstance numberingInstance = new NumberingInstance() { NumberID = numberId };
				AbstractNumId abstractNumId = new AbstractNumId() { Val = numberId };

				numberingInstance.Append(abstractNumId);
				
				context.SaveNumberingDefinition(numberFormat, abstractNum, numberingInstance);
			}
		}
		
		internal DocxUL(IOpenXmlContext context)
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
				InitNumberDefinitions();
				
				foreach (IHtmlNode child in node.Children)
				{
					if (string.Compare(child.Tag, liName, StringComparison.InvariantCultureIgnoreCase) == 0)
					{
						ProcessLi(child, parent);
					}
				}
			}
		}
	}
}

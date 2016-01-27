namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	using DocumentFormat.OpenXml.Packaging;
	
	internal sealed class DocxUL : DocxElement
	{
		private const string elementName = "ul";
		private const string liName = "li";
		
		private void ProcessLi(IHtmlNode li, OpenXmlElement parent)
		{
			OpenXmlElement paragraph = CreateParagraph(li, parent);
			paragraph.Append(GetBulletProperties());
			
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
		
		private ParagraphProperties GetBulletProperties()
		{
			ParagraphProperties paragraphProperties = new ParagraphProperties();
			ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };

			NumberingProperties numberingProperties = new NumberingProperties();
			NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = 0 };
			NumberingId numberingId = new NumberingId() { Val = 1 };
			//Indentation indentation = new Indentation() { Left = (720 * 0).ToString(), Hanging = "360" };

			numberingProperties.Append(numberingLevelReference);
			numberingProperties.Append(numberingId);

			paragraphProperties.Append(paragraphStyleId);
			paragraphProperties.Append(numberingProperties);
			//paragraphProperties.Append(indentation);

			return paragraphProperties;
		}
		
		private void InitNumberDefinitions()
		{
			if (context.MainDocumentPart.NumberingDefinitionsPart == null)
			{
				NumberingDefinitionsPart numberingPart =
					context.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>("numberingDefinitionsPart");
					
				Numbering element = 
					new Numbering(
						new AbstractNum(
							new Level(
								new NumberingFormat() { Val = NumberFormatValues.Bullet },
								new LevelText() { Val = "•" }
							) { LevelIndex = 0 }
						){ AbstractNumberId = 1 },
						new NumberingInstance(
							new AbstractNumId(){ Val = 1 }
						){ NumberID = 1 });

				element.Save(numberingPart);
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
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent)
		{
			if (node == null || !CanConvert(node))
			{
				return;
			}
			
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

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
                Indentation indentation = new Indentation()
                {
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

        private void ProcessLi(DocxNode li, OpenXmlElement parent, NumberFormatValues numberFormat)
        {
            Paragraph paragraph = parent.AppendChild(new Paragraph());
            ParagraphCreated(li, paragraph);

            if (paragraph.ParagraphProperties == null)
            {
                paragraph.ParagraphProperties = new ParagraphProperties();
            }

            SetListProperties(numberFormat, paragraph.ParagraphProperties);

            foreach (DocxNode child in li.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        Run run = paragraph.AppendChild(new Run(new Text()
                        {
                            Text = ClearHtml(child.InnerHtml),
                            Space = SpaceProcessingModeValues.Preserve
                        }));

                        RunCreated(li, run);
                    }
                }
                else
                {
                    child.ParagraphNode = li;
                    child.Parent = parent;
                    li.CopyExtentedStyles(child);
                    ProcessChild(child, ref paragraph);
                }
            }
        }

        private NumberFormatValues GetNumberFormat(DocxNode node)
        {
            NumberFormatValues numberFormat = NumberFormatValues.Decimal;

            string type = node.ExtractAttributeValue("type");

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

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, elementName, StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || !CanConvert(node) || IsHidden(node))
            {
                return;
            }

            paragraph = null;

            if (node.HasChildren)
            {
                NumberFormatValues numberFormat = GetNumberFormat(node);

                InitNumberDefinitions(numberFormat);

                foreach (DocxNode child in node.Children)
                {
                    if (string.Compare(child.Tag, liName, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        node.CopyExtentedStyles(child);
                        ProcessLi(child, node.Parent, numberFormat);
                    }
                }
            }
        }
    }
}

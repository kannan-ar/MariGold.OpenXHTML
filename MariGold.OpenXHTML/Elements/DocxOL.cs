namespace MariGold.OpenXHTML
{
    using System;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxOL : DocxElement
    {
        private const string elementName = "ol";
        private const string liName = "li";

        private bool isParagraphCreated;
        private short numberId;
        
        private void InitNumberDefinitions(NumberFormatValues numberFormat)
        {
            AbstractNum abstractNum = new AbstractNum() { AbstractNumberId = numberId };

            Level level = new Level() { LevelIndex = 0 };
            StartNumberingValue startNumberingValue = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat = new NumberingFormat() { Val = numberFormat };
            LevelText levelText = new LevelText() { Val = "%1." };
            LevelJustification levelJustification = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties = new PreviousParagraphProperties();
            Indentation indentation = new Indentation()
            {
                Left = "720",
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

            context.SaveNumberingDefinition(numberId, abstractNum, numberingInstance);
        }

        private void SetListProperties(ParagraphProperties paragraphProperties)
        {
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };

            NumberingProperties numberingProperties = new NumberingProperties();
            NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId = new NumberingId() { Val = numberId };

            numberingProperties.Append(numberingLevelReference);
            numberingProperties.Append(numberingId);

            paragraphProperties.Append(paragraphStyleId);
            paragraphProperties.Append(numberingProperties);
        }

        private Paragraph CreateParagraph(DocxNode node, OpenXmlElement parent)
        {
            Paragraph para = parent.AppendChild(new Paragraph());
            OnParagraphCreated(node, para);
            OnOLParagraphCreated(this, new ParagraphEventArgs(para));
            return para;
        }

        private void OnOLParagraphCreated(object sender, ParagraphEventArgs args)
        {
            if (!isParagraphCreated)
            {
                if (args.Paragraph.ParagraphProperties == null)
                {
                    args.Paragraph.ParagraphProperties = new ParagraphProperties();
                }

                SetListProperties(args.Paragraph.ParagraphProperties);

                isParagraphCreated = true;
            }
        }

        private void ProcessLi(DocxNode li, OpenXmlElement parent)
        {
            Paragraph paragraph = null;
            isParagraphCreated = false;

            ParagraphCreated = OnOLParagraphCreated;

            foreach (DocxNode child in li.Children)
            {
                if (child.IsText)
                {
                    if (!IsEmptyText(child.InnerHtml))
                    {
                        if (paragraph == null)
                        {
                            paragraph = CreateParagraph(li, parent);
                        }

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
                numberId = ++context.ListNumberId;

                InitNumberDefinitions(numberFormat);

                foreach (DocxNode child in node.Children)
                {
                    if (string.Compare(child.Tag, liName, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        node.CopyExtentedStyles(child);
                        ProcessLi(child, node.Parent);
                    }
                }
            }
        }
    }
}

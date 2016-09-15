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
        private bool isParagraphCreated;

        private Paragraph CreateParagraph(DocxNode node, OpenXmlElement parent)
        {
            Paragraph para = parent.AppendChild(new Paragraph());
            OnParagraphCreated(node, para);
            OnULParagraphCreated(this, new ParagraphEventArgs(para));
            return para;
        }

        private void OnULParagraphCreated(object sender, ParagraphEventArgs args)
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

            ParagraphCreated = OnULParagraphCreated;

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
                Indentation indentation = new Indentation()
                {
                    Start = "720",
                    Hanging = "360"
                };

                previousParagraphProperties.Append(indentation);

                NumberingSymbolRunProperties numberingSymbolRunProperties = new NumberingSymbolRunProperties();
                RunFonts runFonts = new RunFonts()
                {
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
                InitNumberDefinitions();

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

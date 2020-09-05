namespace MariGold.OpenXHTML
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxUL : DocxElement
    {
        private const string elementName = "ul";
        private const string liName = "li";
        private const string levelIndexName = "LevelIndex";
        private const int indent = 360;

        private short gNumberId;
        private bool isParagraphCreated;
        private int gLevelIndex;

        private Paragraph CreateParagraph(DocxNode node, OpenXmlElement parent)
        {
            Paragraph para = parent.AppendChild(new Paragraph());
            OnParagraphCreated(node, para);
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
            else
            {
                DocxParagraphStyle.SetIndentation(args.Paragraph, (gLevelIndex + 1) * indent);
            }
        }

        private void ProcessLi(DocxNode li, OpenXmlElement parent, Dictionary<string, object> properties)
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
                    ProcessChild(child, ref paragraph, properties);
                }
            }
        }

        private void SetListProperties(ParagraphProperties paragraphProperties)
        {
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };

            NumberingProperties numberingProperties = new NumberingProperties();
            NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = gLevelIndex };
            NumberingId numberingId = new NumberingId() { Val = gNumberId };

            numberingProperties.Append(numberingLevelReference);
            numberingProperties.Append(numberingId);

            paragraphProperties.Append(paragraphStyleId);
            paragraphProperties.Append(numberingProperties);
        }

        private void InitNumberDefinitions(int levelIndex)
        {
            AbstractNum abstractNum = new AbstractNum() { AbstractNumberId = gNumberId };

            Level level = new Level() { LevelIndex = levelIndex };
            StartNumberingValue startNumberingValue = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText = new LevelText() { Val = "·" };
            LevelJustification levelJustification = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties = new PreviousParagraphProperties();

            Indentation indentation = new Indentation()
            {
                Start = (indent * (levelIndex + 1)).ToString(),
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

            NumberingInstance numberingInstance = new NumberingInstance() { NumberID = gNumberId };
            AbstractNumId abstractNumId = new AbstractNumId() { Val = gNumberId };

            numberingInstance.Append(abstractNumId);

            context.SaveNumberingDefinition(gNumberId, abstractNum, numberingInstance);
        }

        internal DocxUL(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, elementName, StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (node.IsNull() || !CanConvert(node) || IsHidden(node))
            {
                return;
            }

            paragraph = null;
            int levelIndex = properties.ContainsKey(levelIndexName) ? Convert.ToInt32(properties[levelIndexName]) : 0;

            if (node.HasChildren)
            {
                short numberId = gNumberId = ++context.ListNumberId;

                InitNumberDefinitions(levelIndex);
                
                var newProperties = properties.ToDictionary(x => x.Key, x => x.Value);
                newProperties[levelIndexName] = levelIndex + 1;

                foreach (DocxNode child in node.Children)
                {
                    //Dirty hack to maintain the level since these elements are singleton. When it traverse to inner html elements, the level will increment
                    //then, when it goes back to more root elements, we have to reset the level to its previous state.
                    gLevelIndex = levelIndex;
                    gNumberId = numberId;

                    if (string.Compare(child.Tag, liName, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        node.CopyExtentedStyles(child);
                        ProcessLi(child, node.Parent, newProperties);
                    }
                }
            }
        }
    }
}

namespace MariGold.OpenXHTML
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxOL : DocxElement
    {
        private const string elementName = "ol";
        private const string liName = "li";
        private const string levelIndexName = "LevelIndex";
        private const int indent = 360;
        private const int numIdValue = 1;

        private bool isParagraphCreated;
        private AbstractNum abstractNum;
        private int gLevelIndex;
        private int gLevelId;
        private int gNextLevelId;
        private void DefineLevel(NumberFormatValues numberFormat, int levelIndex)
        {
            Level level = new Level() { LevelIndex = gLevelId };
            StartNumberingValue startNumberingValue = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat = new NumberingFormat() { Val = numberFormat };
            LevelText levelText = new LevelText() { Val = $"%{gLevelId + 1}." }; //Later we need a provison to configure this text.
            LevelJustification levelJustification = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties = new PreviousParagraphProperties();
            Indentation indentation = new Indentation()
            {
                Left = (indent * (levelIndex + 1)).ToString(),
                Hanging = "360"
            };

            previousParagraphProperties.Append(indentation);

            level.Append(startNumberingValue);
            level.Append(numberingFormat);
            level.Append(levelText);
            level.Append(levelJustification);
            level.Append(previousParagraphProperties);

            abstractNum.Append(level);
        }

        private void InitNumberDefinitions()
        {
            abstractNum = new AbstractNum() { AbstractNumberId = numIdValue };

            NumberingInstance numberingInstance = new NumberingInstance() { NumberID = numIdValue };
            AbstractNumId abstractNumId = new AbstractNumId() { Val = numIdValue };

            numberingInstance.Append(abstractNumId);

            context.SaveNumberingDefinition(numIdValue, abstractNum, numberingInstance);
        }

        private void SetListProperties(ParagraphProperties paragraphProperties)
        {
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "ListParagraph" };

            NumberingProperties numberingProperties = new NumberingProperties();
            NumberingLevelReference numberingLevelReference = new NumberingLevelReference() { Val = gLevelId };
            NumberingId numberingId = new NumberingId() { Val = numIdValue };

            numberingProperties.Append(numberingLevelReference);
            numberingProperties.Append(numberingId);

            paragraphProperties.Append(paragraphStyleId);
            paragraphProperties.Append(numberingProperties);
        }

        private Paragraph CreateParagraph(DocxNode node, OpenXmlElement parent)
        {
            Paragraph para = parent.AppendChild(new Paragraph());
            OnParagraphCreated(node, para);
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
            else
            {
                DocxParagraphStyle.SetIndentation(args.Paragraph, (gLevelIndex + 1) * indent);
            }
        }

        private void ProcessLi(DocxNode li, OpenXmlElement parent, Dictionary<string, object> properties)
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
                    ProcessChild(child, ref paragraph, properties);
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

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (node.IsNull() || !CanConvert(node) || IsHidden(node))
            {
                return;
            }

            paragraph = null;
            int levelIndex = properties.ContainsKey(levelIndexName) ? Convert.ToInt32(properties[levelIndexName]) : 0;

            if (abstractNum == null)
            {
                InitNumberDefinitions();
            }

            if (node.HasChildren)
            {
                NumberFormatValues numberFormat = GetNumberFormat(node);

                int levelId = gLevelId = gNextLevelId++;

                DefineLevel(numberFormat, levelIndex);

                var newProperties = properties.ToDictionary(x => x.Key, x => x.Value);
                newProperties[levelIndexName] = levelIndex + 1;

                foreach (DocxNode child in node.Children)
                {
                    gLevelIndex = levelIndex;
                    gLevelId = levelId;

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

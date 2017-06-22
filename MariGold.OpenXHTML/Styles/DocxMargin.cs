namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;

    internal sealed class DocxMargin
    {
        private readonly DocxNode node;
        private const decimal defaultLineHeight = DocxFontStyle.defaultFontSizeInPixel * 20;

        internal const string margin = "margin";
        internal const string marginTop = "margin-top";
        internal const string marginBottom = "margin-bottom";
        internal const string marginLeft = "margin-left";
        internal const string marginRight = "margin-right";
        internal const string lineHeight = "line-height";

        internal DocxMargin(DocxNode node)
        {
            this.node = node;
        }

        internal string GetTopMargin()
        {
            string topMargin = node.ExtractStyleValue(marginTop);

            if (string.IsNullOrEmpty(topMargin))
            {
                topMargin = node.ExtractStyleValue(margin);
            }

            return topMargin;
        }

        internal string GetBottomMargin()
        {
            string bottomMargin = node.ExtractStyleValue(marginBottom);

            if (string.IsNullOrEmpty(bottomMargin))
            {
                bottomMargin = node.ExtractStyleValue(margin);
            }

            return bottomMargin;
        }

        internal string GetLeftMargin()
        {
            string leftMargin = node.ExtractStyleValue(marginLeft);

            if (string.IsNullOrEmpty(leftMargin))
            {
                leftMargin = node.ExtractStyleValue(margin);
            }

            return leftMargin;
        }

        internal string GetRightMargin()
        {
            string rightMargin = node.ExtractStyleValue(marginRight);

            if (string.IsNullOrEmpty(rightMargin))
            {
                rightMargin = node.ExtractStyleValue(margin);
            }

            return rightMargin;
        }

        internal void SetLeftMargin(string value)
        {
            node.SetExtentedStyle(marginLeft, value);
        }

        internal void ProcessParagraphMargin(ParagraphProperties properties)
        {
            string topMargin = GetTopMargin();
            string bottomMargin = GetBottomMargin();
            string leftMargin = GetLeftMargin();
            string rightMargin = GetRightMargin();
            string line = node.ExtractStyleValue(lineHeight);

            if (!string.IsNullOrEmpty(topMargin) || !string.IsNullOrEmpty(bottomMargin) || !string.IsNullOrEmpty(line))
            {
                SpacingBetweenLines spacing = new SpacingBetweenLines();

                if (!string.IsNullOrEmpty(topMargin))
                {
                    decimal dxa = DocxUnits.GetDxaFromStyle(topMargin);

                    if (dxa != -1)
                    {
                        spacing.Before = dxa.ToString();
                    }
                }

                if (!string.IsNullOrEmpty(bottomMargin))
                {
                    decimal dxa = DocxUnits.GetDxaFromStyle(bottomMargin);

                    if (dxa != -1)
                    {
                        spacing.After = decimal.Round(dxa).ToString();
                    }
                }

                if (!string.IsNullOrEmpty(line) && !line.CompareStringInvariantCultureIgnoreCase(DocxFontStyle.normal))
                {
                    decimal number;
                    decimal dxa = -1;
                    LineSpacingRuleValues lineSpacingRuleValues = LineSpacingRuleValues.AtLeast;

                    if (decimal.TryParse(line, out number))
                    {
                        dxa = DocxUnits.GetDxaFromNumber(number);
                        dxa = dxa - defaultLineHeight;//Removing the default line height
                    }
                    else if (line.Contains("%"))
                    {
                        line = line.Replace("%", string.Empty);

                        if (decimal.TryParse(line, out number))
                        {
                            dxa = (number / 100) * DocxFontStyle.defaultFontSizeInPixel;
                            dxa = dxa - defaultLineHeight;//Removing the default line height
                        }
                    }
                    else
                    {
                        dxa = DocxUnits.GetDxaFromStyle(line);
                        //lineSpacingRuleValues = LineSpacingRuleValues.Exact;
                    }

                    dxa = decimal.Round(dxa);

                    if (dxa > 0)
                    {
                        spacing.LineRule = lineSpacingRuleValues;
                        spacing.Line = dxa.ToString();
                    }
                }

                if (spacing.HasAttributes)
                {
                    properties.Append(spacing);
                }
            }

            if (!string.IsNullOrEmpty(leftMargin) || !string.IsNullOrEmpty(rightMargin))
            {
                Indentation ind = new Indentation();

                if (!string.IsNullOrEmpty(leftMargin))
                {
                    decimal dxa = DocxUnits.GetDxaFromStyle(leftMargin);

                    if (dxa != -1)
                    {
                        ind.Left = dxa.ToString();
                    }
                }

                if (!string.IsNullOrEmpty(rightMargin))
                {
                    decimal dxa = DocxUnits.GetDxaFromStyle(rightMargin);

                    if (dxa != -1)
                    {
                        ind.Right = dxa.ToString();
                    }
                }

                if (ind.HasAttributes)
                {
                    properties.Append(ind);
                }
            }
        }

        internal static void SetTopMargin(string style, ParagraphProperties properties)
        {
            decimal dxa = DocxUnits.GetDxaFromStyle(style);

            if (dxa != -1)
            {
                SpacingBetweenLines spacing = new SpacingBetweenLines();

                spacing.Before = dxa.ToString();
                properties.Append(spacing);
            }
        }

        internal static void SetBottomMargin(string style, ParagraphProperties properties)
        {
            decimal dxa = DocxUnits.GetDxaFromStyle(style);

            if (dxa != -1)
            {
                SpacingBetweenLines spacing = new SpacingBetweenLines();

                spacing.After = dxa.ToString();
                properties.Append(spacing);
            }
        }
    }
}

namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;

    internal sealed class DocxTableCellStyle
    {
        private const string colspan = "colspan";

        private void ProcessBorders(DocxNode node, DocxTableProperties docxProperties,
            TableCell cell)
        {
            string borderStyle = node.ExtractStyleValue(DocxBorder.borderName);
            string leftBorder = node.ExtractStyleValue(DocxBorder.leftBorderName);
            string topBorder = node.ExtractStyleValue(DocxBorder.topBorderName);
            string rightBorder = node.ExtractStyleValue(DocxBorder.rightBorderName);
            string bottomBorder = node.ExtractStyleValue(DocxBorder.bottomBorderName);

            TableCellBorders cellBorders = new TableCellBorders();

            DocxBorder.ApplyBorders(cellBorders, borderStyle, leftBorder, topBorder,
                rightBorder, bottomBorder, docxProperties.HasDefaultBorder);

            if (cellBorders.HasChildren)
            {
                AssignTableCellPropertiesIfEmpty(cell);
                cell.TableCellProperties.Append(cellBorders);
            }
        }

        private void ProcessColSpan(DocxNode node, TableCell cell)
        {
            if (Int32.TryParse(node.ExtractAttributeValue(colspan), out int value))
            {
                if (value > 1)
                {
                    AssignTableCellPropertiesIfEmpty(cell);
                    cell.TableCellProperties.Append(new GridSpan() { Val = value });
                }
            }
        }

        private void ProcessWidth(DocxNode node, TableCell cell)
        {
            string width = node.ExtractStyleValue(DocxUnits.width);

            if (!string.IsNullOrEmpty(width))
            {
                if (DocxUnits.TableUnitsFromStyle(width, out decimal value, out TableWidthUnitValues unit))
                {
                    TableCellWidth cellWidth = new TableCellWidth()
                    {
                        Width = value.ToString(),
                        Type = unit
                    };

                    AssignTableCellPropertiesIfEmpty(cell);
                    cell.TableCellProperties.Append(cellWidth);
                }
            }
        }

        private void ProcessVerticalAlignment(DocxNode node, TableCell cell)
        {
            string alignment = node.ExtractStyleValue(DocxAlignment.verticalAlign);

            if (!string.IsNullOrEmpty(alignment))
            {
                if (DocxAlignment.GetCellVerticalAlignment(alignment, out TableVerticalAlignmentValues value))
                {
                    AssignTableCellPropertiesIfEmpty(cell);
                    cell.TableCellProperties.Append(new TableCellVerticalAlignment() { Val = value });
                }
            }
        }

        private void AssignTableCellPropertiesIfEmpty(TableCell cell)
        {
            if (cell.TableCellProperties == null)
            {
                cell.TableCellProperties = new TableCellProperties();
            }
        }

        internal bool HasRowSpan { get; set; }

        internal void Process(TableCell cell, DocxTableProperties docxProperties, DocxNode node)
        {
            ProcessColSpan(node, cell);
            ProcessWidth(node, cell);

            if (HasRowSpan)
            {
                AssignTableCellPropertiesIfEmpty(cell);
                cell.TableCellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
            }

            //Processing border should be after colspan
            ProcessBorders(node, docxProperties, cell);

            string backgroundColor = node.ExtractStyleValue(DocxColor.backGroundColor);

            if (!string.IsNullOrEmpty(backgroundColor))
            {
                AssignTableCellPropertiesIfEmpty(cell);
                DocxColor.ApplyBackGroundColor(backgroundColor, cell.TableCellProperties);
            }

            ProcessVerticalAlignment(node, cell);
        }

        internal static DocxNode GetHtmlNodeForTableCellContent(DocxNode node)
        {
            node.RemoveStyles(DocxBorder.borderName, DocxBorder.leftBorderName, DocxBorder.rightBorderName,
                DocxBorder.topBorderName, DocxBorder.bottomBorderName, DocxMargin.margin, DocxMargin.marginLeft,
                DocxMargin.marginRight, DocxMargin.marginTop, DocxMargin.marginBottom, DocxColor.backGroundColor);

            return node;
        }
    }
}

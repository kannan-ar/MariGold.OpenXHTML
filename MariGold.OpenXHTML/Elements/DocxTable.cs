namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.Linq;

    internal sealed class DocxTable : DocxElement
    {
        private void SetThStyleToRun(IHtmlNode run)
        {
            DocxNode docxNode = new DocxNode(run);

            string value = docxNode.ExtractStyleValue(DocxFont.fontWeight);

            if (string.IsNullOrEmpty(value))
            {
                docxNode.SetStyleValue(DocxFont.fontWeight, DocxFont.bold);
            }
        }

        private void ProcessTd(int colIndex, IHtmlNode td, TableRow row, DocxTableProperties tableProperties)
        {
            TableCell cell = new TableCell();
            bool hasRowSpan = false;

            DocxNode docxNode = new DocxNode(td);
            string rowSpan = docxNode.ExtractAttributeValue(DocxTableProperties.rowSpan);
            Int32 rowSpanValue;
            if (Int32.TryParse(rowSpan, out rowSpanValue))
            {
                tableProperties.RowSpanInfo[colIndex] = rowSpanValue - 1;
                hasRowSpan = true;
            }

            DocxTableCellStyle style = new DocxTableCellStyle();
            style.HasRowSpan = hasRowSpan;
            style.Process(cell, tableProperties, td);

            if (td.HasChildren)
            {
                Paragraph para = null;

                foreach (IHtmlNode child in td.Children)
                {
                    //If the cell is th header, apply font-weight:bold to the text
                    if (tableProperties.IsCellHeader)
                    {
                        SetThStyleToRun(child);
                    }

                    if (child.IsText)
                    {
                        if (!IsEmptyText(child.InnerHtml))
                        {
                            if (para == null)
                            {
                                para = cell.AppendChild(new Paragraph());
                                ParagraphCreated(DocxTableCellStyle.GetHtmlNodeForTableCellContent(td), para);
                            }

                            Run run = para.AppendChild(new Run(new Text()
                            {
                                Text = ClearHtml(child.InnerHtml),
                                Space = SpaceProcessingModeValues.Preserve
                            }));

                            RunCreated(child, run);
                        }
                    }
                    else
                    {
                        ProcessChild(new DocxProperties(child, DocxTableCellStyle.GetHtmlNodeForTableCellContent(td), cell), ref para);
                    }
                }
            }

            //The last element of the table cell must be a paragraph.
            var lastElement = cell.Elements().LastOrDefault();

            if (lastElement == null || !(lastElement is Paragraph))
            {
                cell.AppendChild(new Paragraph());
            }

            row.Append(cell);
        }

        private void ProcessVerticalSpan(ref int colIndex, TableRow row, DocxTableProperties docxProperties)
        {
            int rowSpan;

            docxProperties.RowSpanInfo.TryGetValue(colIndex, out rowSpan);

            while (rowSpan > 0)
            {
                TableCell cell = new TableCell();

                cell.TableCellProperties = new TableCellProperties();
                cell.TableCellProperties.Append(new VerticalMerge());

                cell.AppendChild(new Paragraph());

                row.Append(cell);

                docxProperties.RowSpanInfo[colIndex] = --rowSpan;
                ++colIndex;
                docxProperties.RowSpanInfo.TryGetValue(colIndex, out rowSpan);
            }
        }

        private void ProcessTr(IHtmlNode tr, Table table, DocxTableProperties tableProperties)
        {
            if (tr.HasChildren)
            {
                TableRow row = new TableRow();
                DocxNode trNode = new DocxNode(tr);

                DocxTableRowStyle style = new DocxTableRowStyle();
                style.Process(row, tableProperties);

                int colIndex = 0;

                foreach (IHtmlNode td in tr.Children)
                {
                    ProcessVerticalSpan(ref colIndex, row, tableProperties);

                    tableProperties.IsCellHeader = string.Compare(td.Tag, DocxTableProperties.thName, StringComparison.InvariantCultureIgnoreCase) == 0;

                    if (string.Compare(td.Tag, DocxTableProperties.tdName, StringComparison.InvariantCultureIgnoreCase) == 0 || tableProperties.IsCellHeader)
                    {
                        trNode.CopyStyles(td, DocxColor.backGroundColor);
                        ProcessTd(colIndex++, td, row, tableProperties);
                    }
                }

                if (colIndex < tableProperties.RowSpanInfo.Count)
                {
                    ProcessVerticalSpan(ref colIndex, row, tableProperties);
                }

                table.Append(row);
            }
        }

        internal DocxTable(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(IHtmlNode node)
        {
            return string.Compare(node.Tag, DocxTableProperties.tableName, StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxProperties properties, ref Paragraph paragraph)
        {
            if (properties.CurrentNode == null || properties.Parent == null || !CanConvert(properties.CurrentNode))
            {
                return;
            }

            paragraph = null;

            if (properties.CurrentNode.HasChildren)
            {
                Table table = new Table();
                DocxTableProperties tableProperties = new DocxTableProperties();

                tableProperties.FetchTableProperties(properties.CurrentNode);
                tableProperties.ApplyTableProperties(table, properties.CurrentNode);

                foreach (IHtmlNode tr in properties.CurrentNode.Children)
                {
                    if (string.Compare(tr.Tag, DocxTableProperties.trName, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        ProcessTr(tr, table, tableProperties);
                    }
                }

                properties.Parent.Append(table);
            }
        }
    }
}

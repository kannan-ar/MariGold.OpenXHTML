namespace MariGold.OpenXHTML
{
    using System;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System.Linq;

    internal sealed class DocxTable : DocxElement
    {
        private void SetThStyleToRun(DocxNode run)
        {
            string value = run.ExtractStyleValue(DocxFontStyle.fontWeight);

            if (string.IsNullOrEmpty(value))
            {
                run.SetExtentedStyle(DocxFontStyle.fontWeight, DocxFontStyle.bold);
            }
        }

        private void ProcessTd(int colIndex, DocxNode td, TableRow row, DocxTableProperties tableProperties)
        {
            TableCell cell = new TableCell();
            bool hasRowSpan = false;

            string rowSpan = td.ExtractAttributeValue(DocxTableProperties.rowSpan);
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

                //If the cell is th header, apply font-weight:bold to the text
                if (tableProperties.IsCellHeader)
                {
                    SetThStyleToRun(td);
                }

                foreach (DocxNode child in td.Children)
                {
                    td.CopyExtentedStyles(child);
                    
                    if (child.IsText)
                    {
                        if (!IsEmptyText(child.InnerHtml))
                        {
                            if (para == null)
                            {
                                para = cell.AppendChild(new Paragraph());
                                OnParagraphCreated(DocxTableCellStyle.GetHtmlNodeForTableCellContent(td), para);
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
                        child.ParagraphNode = DocxTableCellStyle.GetHtmlNodeForTableCellContent(td);
                        child.Parent = cell;
                        td.CopyExtentedStyles(child);
                        ProcessChild(child, ref para);
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

        private void ProcessTr(DocxNode tr, Table table, DocxTableProperties tableProperties)
        {
            if (tr.HasChildren)
            {
                TableRow row = new TableRow();

                DocxTableRowStyle style = new DocxTableRowStyle();
                style.Process(row, tableProperties);

                int colIndex = 0;

                foreach (DocxNode td in tr.Children)
                {
                    ProcessVerticalSpan(ref colIndex, row, tableProperties);

                    tableProperties.IsCellHeader = string.Compare(td.Tag, DocxTableProperties.thName, StringComparison.InvariantCultureIgnoreCase) == 0;

                    if (string.Compare(td.Tag, DocxTableProperties.tdName, StringComparison.InvariantCultureIgnoreCase) == 0 || tableProperties.IsCellHeader)
                    {
                        tr.CopyExtentedStyles(td);
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

        private void ProcessTBody(DocxNode tbody, Table table, DocxTableProperties tableProperties)
        {
            foreach (DocxNode tr in tbody.Children)
            {
                if (string.Compare(tr.Tag, DocxTableProperties.trName, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    tbody.CopyExtentedStyles(tr);
                    ProcessTr(tr, table, tableProperties);
                }
            }
        }

        internal DocxTable(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Compare(node.Tag, DocxTableProperties.tableName, StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (node.IsNull() || node.Parent == null || !CanConvert(node) || IsHidden(node))
            {
                return;
            }

            paragraph = null;

            if (node.HasChildren)
            {
                Table table = new Table();
                DocxTableProperties tableProperties = new DocxTableProperties();

                tableProperties.FetchTableProperties(node);
                tableProperties.ApplyTableProperties(table, node);

                foreach (DocxNode child in node.Children)
                {
                    if (string.Compare(child.Tag, DocxTableProperties.trName, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        node.CopyExtentedStyles(child);
                        ProcessTr(child, table, tableProperties);
                    }
                    else if (string.Compare(child.Tag, DocxTableProperties.tbody, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        node.CopyExtentedStyles(child);
                        ProcessTBody(child, table, tableProperties);
                    }
                }

                node.Parent.Append(table);
            }
        }
    }
}

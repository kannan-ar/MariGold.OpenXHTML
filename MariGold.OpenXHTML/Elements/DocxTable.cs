namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using MariGold.OpenXHTML.Styles;
    using System;
    using System.Collections.Generic;
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

        private void ProcessTd(int colIndex, DocxNode td, TableRow row, DocxTableProperties tableProperties, Dictionary<string, object> properties)
        {
            TableCell cell = new TableCell();
            bool hasRowSpan = false;

            string rowSpan = td.ExtractAttributeValue(DocxTableProperties.rowSpan);
            if (int.TryParse(rowSpan, out int rowSpanValue))
            {
                tableProperties.UpdateRowSpan(colIndex, rowSpanValue - 1, td);
                hasRowSpan = true;
            }

            DocxTableCellStyle style = new DocxTableCellStyle
            {
                HasRowSpan = hasRowSpan
            };

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
                    td.ApplyBlockStyles(child);

                    if (child.IsText)
                    {
                        if (!IsEmptyText(child.InnerHtml))
                        {
                            if (para == null)
                            {
                                para = cell.AppendChild(new Paragraph());
                                OnParagraphCreated(DocxTableCellStyle.GetHtmlNodeForTableCellContent(td.Clone()), para);
                            }

                            Run run = para.AppendChild(new Run(new[] { new Text()
                            {
                                Text = ClearHtml(child.InnerHtml),
                                Space = SpaceProcessingModeValues.Preserve
                            }}));

                            RunCreated(child, run);
                        }
                    }
                    else
                    {
                        child.ParagraphNode = DocxTableCellStyle.GetHtmlNodeForTableCellContent(td.Clone());
                        child.Parent = cell;
                        td.CopyExtentedStyles(child);
                        td.ApplyBlockStyles(child);
                        ProcessChild(child, ref para, properties);
                    }
                }
            }

            //The last element of the table cell must be a paragraph.
            var lastElement = cell.Elements().LastOrDefault();

            if (!(lastElement is Paragraph))
            {
                cell.AppendChild(new Paragraph());
            }

            row.Append(new[] { cell });
        }

        private void ProcessVerticalSpan(ref int colIndex, TableRow row, DocxTableProperties docxProperties)
        {
            var hasRowSpan = docxProperties.TryGetRowSpan(colIndex, out int rowSpan, out DocxNode node);

            while (hasRowSpan)
            {
                DocxTableCellStyle style = new DocxTableCellStyle();
                TableCell cell = new TableCell
                {
                    TableCellProperties = new TableCellProperties()
                };

                cell.TableCellProperties.Append(new[] { new VerticalMerge { Val = MergedCellValues.Continue } } );

                style.Process(cell, docxProperties, node);

                cell.AppendChild(new Paragraph());

                row.Append(new[] { cell } );

                docxProperties.UpdateRowSpan(colIndex, rowSpan - 1, node);
                ++colIndex;
                hasRowSpan = docxProperties.TryGetRowSpan(colIndex, out rowSpan, out node);
            }
        }

        private void ProcessTr(DocxNode tr, Table table, DocxTableProperties tableProperties, Dictionary<string, object> properties)
        {
            if (tr.HasChildren)
            {
                TableRow row = new TableRow();

                DocxTableRowStyle style = new DocxTableRowStyle();
                style.Process(row, tableProperties);

                int colIndex = 0;

                foreach (DocxNode td in tr.Children)
                {
                    tableProperties.IsCellHeader = string.Compare(td.Tag, DocxTableProperties.thName, StringComparison.InvariantCultureIgnoreCase) == 0;

                    if (string.Compare(td.Tag, DocxTableProperties.tdName, StringComparison.InvariantCultureIgnoreCase) == 0 || tableProperties.IsCellHeader)
                    {
                        ProcessVerticalSpan(ref colIndex, row, tableProperties);

                        tr.CopyExtentedStyles(td);
                        ProcessTd(colIndex++, td, row, tableProperties, properties);
                    }
                }

                if (colIndex < tableProperties.GetRowSpanCount())
                {
                    ProcessVerticalSpan(ref colIndex, row, tableProperties);
                }

                table.Append(new[] { row } );
            }
        }

        private void ProcessGroupElement(DocxNode tbody, Table table, DocxTableProperties tableProperties, Dictionary<string, object> properties)
        {
            foreach (DocxNode tr in tbody.Children)
            {
                if (string.Compare(tr.Tag, DocxTableProperties.trName, StringComparison.InvariantCultureIgnoreCase) == 0)
                {
                    tbody.CopyExtentedStyles(tr);
                    ProcessTr(tr, table, tableProperties, properties);
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

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
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
                        ProcessTr(child, table, tableProperties, properties);
                    }
                    else if (tableProperties.IsGroupElement(child.Tag))
                    {
                        node.CopyExtentedStyles(child);
                        ProcessGroupElement(child, table, tableProperties, properties);
                    }
                }

                node.Parent.Append(new[] { table } );
            }
        }
    }
}

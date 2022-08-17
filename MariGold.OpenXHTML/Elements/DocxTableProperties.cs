namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections.Generic;

    internal sealed class DocxTableProperties
    {
        private bool hasDefaultHeader;
        private bool isCellHeader;
        private short? cellPadding;
        private short? cellSpacing;
        private Dictionary<int, RowSpan> rowSpanInfo;

        private sealed class RowSpan
        {
            internal RowSpan(int count, DocxNode node)
            {
                Count = count;
                Node = node;
            }

            internal int Count { get; set; }
            internal DocxNode Node { get; set; }
        }

        internal const string tableName = "table";
        internal const string thead = "thead";
        internal const string tbody = "tbody";
        internal const string tfoot = "tfoot";
        internal const string trName = "tr";
        internal const string tdName = "td";
        internal const string thName = "th";
        internal const string tableGridName = "TableGrid";
        internal const string cellSpacingName = "cellspacing";
        internal const string cellPaddingName = "cellpadding";
        internal const string colspan = "colspan";
        internal const string rowSpan = "rowspan";

        internal bool HasDefaultBorder
        {
            get
            {
                return hasDefaultHeader;
            }

            set
            {
                hasDefaultHeader = value;
            }
        }

        internal bool IsCellHeader
        {
            get
            {
                return isCellHeader;
            }

            set
            {
                isCellHeader = value;
            }
        }

        internal short? CellPadding
        {
            get
            {
                return cellPadding;
            }

            set
            {
                cellPadding = value;
            }
        }

        internal short? CellSpacing
        {
            get
            {
                return cellSpacing;
            }

            set
            {
                cellSpacing = value;
            }
        }

        private int GetTdCount(DocxNode element)
        {
            int count = 0;

            if (element != null && element.HasChildren)
            {
                foreach (DocxNode child in element.Children)
                {
                    if (IsGroupElement(child.Tag))
                    {
                        return GetTdCount(child);
                    }
                    else if (string.Compare(child.Tag, DocxTableProperties.trName, StringComparison.InvariantCultureIgnoreCase) == 0)
                    {
                        foreach (DocxNode td in child.Children)
                        {
                            if (string.Compare(td.Tag, DocxTableProperties.tdName, StringComparison.InvariantCultureIgnoreCase) == 0 ||
                                string.Compare(td.Tag, DocxTableProperties.thName, StringComparison.InvariantCultureIgnoreCase) == 0)
                            {
                                string colSpan = td.ExtractAttributeValue("colspan");

                                if (!string.IsNullOrEmpty(colSpan) && Int32.TryParse(colSpan, out int colspanValue))
                                {
                                    count += colspanValue;
                                }
                                else
                                {
                                    ++count;
                                }
                            }
                        }

                        //Counted first row's td count. Thus exiting
                        break;
                    }
                }
            }

            return count;
        }

        internal void FetchTableProperties(DocxNode node)
        {
            this.HasDefaultBorder = node.ExtractAttributeValue(DocxBorder.borderName) == "1";

            if (short.TryParse(node.ExtractAttributeValue(cellSpacingName), out short value))
            {
                this.CellSpacing = value;
            }

            if (short.TryParse(node.ExtractAttributeValue(cellPaddingName), out value))
            {
                this.CellPadding = value;
            }
        }

        internal void ApplyTableProperties(Table table, DocxNode node)
        {
            TableProperties tableProp = new TableProperties();

            TableStyle tableStyle = new TableStyle() { Val = DocxTableProperties.tableGridName };

            tableProp.Append(tableStyle);

            DocxTableStyle style = new DocxTableStyle();
            style.Process(tableProp, this, node);

            table.AppendChild(tableProp);

            int count = GetTdCount(node);

            rowSpanInfo = new Dictionary<int, RowSpan>();

            if (count > 0)
            {
                TableGrid tg = new TableGrid();

                for (int i = 0; i < count; i++)
                {
                    rowSpanInfo.Add(i, new RowSpan(0, null));
                    tg.AppendChild(new GridColumn());
                }

                table.AppendChild(tg);
            }
        }

        internal bool IsGroupElement(string tag)
        {
            return string.Compare(tag, DocxTableProperties.thead, StringComparison.InvariantCultureIgnoreCase) == 0 ||
                        string.Compare(tag, DocxTableProperties.tbody, StringComparison.InvariantCultureIgnoreCase) == 0 ||
                        string.Compare(tag, DocxTableProperties.tfoot, StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal void UpdateRowSpan(int colIndex, int count, DocxNode node)
        {
            rowSpanInfo[colIndex].Count = count;
            rowSpanInfo[colIndex].Node = node;
        }

        internal bool TryGetRowSpan(int colIndex, out int rowSpan, out DocxNode node)
        {
            rowSpan = 0;
            node = null;

            var recordFound = rowSpanInfo.TryGetValue(colIndex, out RowSpan row);

            if(recordFound)
            {
                rowSpan = row.Count;
                node = row.Node;
            }

            return rowSpan > 0;
        }

        internal int GetRowSpanCount()
        {
            return rowSpanInfo.Count;
        }
    }
}

namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.Collections.Generic;
	using System.Linq;
	
	internal sealed class DocxTableProperties
	{
		private bool hasDefaultHeader;
		private bool isCellHeader;
		private Int16? cellPadding;
		private Int16? cellSpacing;
		private Dictionary<int, int> rowSpanInfo;
		
		internal const string tableName = "table";
        internal const string tbody = "tbody";
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
		
		internal Int16? CellPadding
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
		
		internal Int16? CellSpacing
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
		
		internal Dictionary<int, int> RowSpanInfo
		{
			get
			{
				return rowSpanInfo;
			}
		}

        private Int32 GetTdCount(DocxNode table)
		{
			int count = 0;
			
			if (table != null && table.HasChildren)
			{
				foreach (DocxNode tr in table.Children)
				{
					if (string.Compare(tr.Tag, DocxTableProperties.trName, StringComparison.InvariantCultureIgnoreCase) == 0)
					{
						foreach (DocxNode td in tr.Children)
						{
							if (string.Compare(td.Tag, DocxTableProperties.tdName, StringComparison.InvariantCultureIgnoreCase) == 0 ||
							    string.Compare(td.Tag, DocxTableProperties.thName, StringComparison.InvariantCultureIgnoreCase) == 0)
							{
								string colSpan = table.ExtractAttributeValue("colspan");
								Int32 colspanValue;
								
								if (!string.IsNullOrEmpty(colspan) && Int32.TryParse(colspan, out colspanValue))
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
			
			Int16 value;

            if (Int16.TryParse(node.ExtractAttributeValue(cellSpacingName), out value))
			{
				this.CellSpacing = value;
			}

            if (Int16.TryParse(node.ExtractAttributeValue(cellPaddingName), out value))
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
			
			rowSpanInfo = new Dictionary<int, int>();
			
			if (count > 0)
			{
				TableGrid tg = new TableGrid();
				
				for (int i = 0; i < count; i++)
				{
					rowSpanInfo.Add(i, 0);
					tg.AppendChild(new GridColumn());
				}
				
				table.AppendChild(tg);
			}
		}
	}
}

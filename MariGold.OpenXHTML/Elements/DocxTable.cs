namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.Linq;
	
	internal sealed class DocxTable : DocxElement
	{
		private const string tableName = "table";
		private const string trName = "tr";
		private const string tdName = "td";
		private const string thName = "th";
		private const string tableGridName = "TableGrid";
		private const string cellSpacingName = "cellspacing";
		private const string cellPaddingName = "cellpadding";
		
		private DocxTableProperties GetTableProperties(IHtmlNode node)
		{
			DocxNode docxNode = new DocxNode(node);
			DocxTableProperties docxProperties = new DocxTableProperties();
				
			docxProperties.HasDefaultBorder = docxNode.ExtractAttributeValue(DocxBorder.borderName) == "1";
			
			Int16 value;
			
			if (Int16.TryParse(docxNode.ExtractAttributeValue(cellSpacingName), out value))
			{
				docxProperties.CellSpacing = value;
			}
			
			if (Int16.TryParse(docxNode.ExtractAttributeValue(cellPaddingName), out value))
			{
				docxProperties.CellPadding = value;
			}
			
			return docxProperties;
		}
		
		private void ProcessTd(IHtmlNode td, TableRow row, DocxTableProperties docxProperties)
		{
			if (td.HasChildren)
			{
				TableCell cell = new TableCell();
				
				DocxTableStyle style = new DocxTableStyle();
				style.Process(cell, docxProperties, td);
				
				Paragraph para = null;
				
				foreach (IHtmlNode child in td.Children)
				{
					if (child.IsText)
					{
						if (para == null)
						{
							para = new Paragraph();
							ParagraphCreated(td, para);
						}
						
						Run run = para.AppendChild(new Run(new Text(child.InnerHtml)));
						RunCreated(child, run);
					}
					else
					{
						if (para != null)
						{
							cell.Append(para);
						}
						
						ProcessChild(child, cell, ref para);
					}
				}
				
				if (para != null)
				{
					cell.Append(para);
				}
				
				row.Append(cell);
			}
		}
		
		private void ProcessTr(IHtmlNode tr, Table table, DocxTableProperties docxProperties)
		{
			if (tr.HasChildren)
			{
				TableRow row = new TableRow();
				
				DocxTableStyle style = new DocxTableStyle();
				style.Process(row, docxProperties, tr);
			
				foreach (IHtmlNode td in tr.Children)
				{
					docxProperties.IsCellHeader = string.Compare(td.Tag, thName, StringComparison.InvariantCultureIgnoreCase) == 0;
					
					if (string.Compare(td.Tag, tdName, StringComparison.InvariantCultureIgnoreCase) == 0 || docxProperties.IsCellHeader)
					{
						ProcessTd(td, row, docxProperties);
					}
				}
				
				table.Append(row);
			}
		}
		
		private void ApplyTableProperties(Table table, DocxTableProperties docxProperties, IHtmlNode node)
		{
			TableProperties tableProp = new TableProperties();
			
			TableStyle tableStyle = new TableStyle() { Val = tableGridName };
			
			tableProp.Append(tableStyle);
			
			DocxTableStyle style = new DocxTableStyle();
			style.Process(tableProp, docxProperties, node);
			
			table.AppendChild(tableProp);
			
			int count = node.Children.Count();
			
			if (count > 0)
			{
				TableGrid tg = new TableGrid();
				
				for (int i = 0; i < count; i++)
				{
					tg.AppendChild(new GridColumn());
				}
				
				table.AppendChild(tg);
			}
		}
		
		internal DocxTable(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, tableName, StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node == null || parent == null || !CanConvert(node))
			{
				return;
			}
			
			paragraph = null;
			
			if (node.HasChildren)
			{
				Table table = new Table();
				DocxTableProperties docxProperties = GetTableProperties(node);
				
				ApplyTableProperties(table, docxProperties, node);
				
				foreach (IHtmlNode tr in node.Children)
				{
					if (string.Compare(tr.Tag, trName, StringComparison.InvariantCultureIgnoreCase) == 0)
					{
						ProcessTr(tr, table, docxProperties);
					}
				}
				
				parent.Append(table);
			}
		}
	}
}

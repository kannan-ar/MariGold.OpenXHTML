namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.Linq;
	using System.Collections.Generic;
	
	internal sealed class DocxTable : DocxElement
	{
		private void ApplyTableCellProperties(Table table)
		{
			
		}
		
		private void ProcessTd(IHtmlNode td, TableRow row, bool isHeader)
		{
			if (td.HasChildren)
			{
				TableCell cell = new TableCell();
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
		
		private void ProcessTr(IHtmlNode tr, Table table)
		{
			if (tr.HasChildren)
			{
				TableRow row = new TableRow();
				
				foreach (IHtmlNode td in tr.Children)
				{
					bool isHeader = string.Compare(td.Tag, "th", StringComparison.InvariantCultureIgnoreCase) == 0;
					
					if (string.Compare(td.Tag, "td", StringComparison.InvariantCultureIgnoreCase) == 0 || isHeader)
					{
						ProcessTd(td, row, isHeader);
					}
				}
				
				table.Append(row);
			}
		}
		
		private void ApplyTableBorder(TableStyle tableStyle, IHtmlNode node)
		{
			string borderStyle = ExtractAttributeValue("border", node);
			
			if (borderStyle == "1")
			{
				TableBorders tableBorders = new TableBorders();
				
				DocxBorder.ApplyDefaultBorders(tableBorders);
			
				tableStyle.Append(tableBorders);
			}
			else
			{
				borderStyle = ExtractStyleValue("border", node);
				
				if (!string.IsNullOrEmpty(borderStyle))
				{
					
				}
			}
		}
		
		private void ApplyTableProperties(Table table, IHtmlNode node)
		{
			TableProperties tableProp = new TableProperties();
			
			TableStyle tableStyle = new TableStyle() { Val = "TableGrid" };
			tableProp.Append(tableStyle);
			
			ApplyTableBorder(tableStyle, node);
			
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
			return string.Compare(node.Tag, "table", StringComparison.InvariantCultureIgnoreCase) == 0;
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
				
				ApplyTableProperties(table, node);
				
				foreach (IHtmlNode tr in node.Children)
				{
					if (string.Compare(tr.Tag, "tr", StringComparison.InvariantCultureIgnoreCase) == 0)
					{
						ProcessTr(tr, table);
					}
				}
				
				parent.Append(table);
			}
		}
	}
}

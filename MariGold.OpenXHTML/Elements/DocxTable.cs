namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
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
		
		private void ProcessTd(IHtmlNode td, TableRow row, DocxTableProperties docxProperties)
		{
			if (td.HasChildren)
			{
				TableCell cell = new TableCell();
				
				DocxTableCellStyle style = new DocxTableCellStyle();
				style.Process(cell, docxProperties, td);
				
				Paragraph para = null;
				
				foreach (IHtmlNode child in td.Children)
				{
					//If the cell is th header, apply font-weight:bold to the text
					if (docxProperties.IsCellHeader)
					{
						SetThStyleToRun(child);
					}
					
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
				
				DocxTableRowStyle style = new DocxTableRowStyle();
				style.Process(row, docxProperties);
			
				foreach (IHtmlNode td in tr.Children)
				{
					docxProperties.IsCellHeader = string.Compare(td.Tag, DocxTableProperties.thName, StringComparison.InvariantCultureIgnoreCase) == 0;
					
					if (string.Compare(td.Tag, DocxTableProperties.tdName, StringComparison.InvariantCultureIgnoreCase) == 0 || docxProperties.IsCellHeader)
					{
						ProcessTd(td, row, docxProperties);
					}
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
				DocxTableProperties docxProperties = new DocxTableProperties();
				
				docxProperties.FetchTableProperties(node);
				docxProperties.ApplyTableProperties(table, node);
				
				foreach (IHtmlNode tr in node.Children)
				{
					if (string.Compare(tr.Tag, DocxTableProperties.trName, StringComparison.InvariantCultureIgnoreCase) == 0)
					{
						ProcessTr(tr, table, docxProperties);
					}
				}
				
				parent.Append(table);
			}
		}
	}
}

namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxTable : DocxElement
	{
		private void ProcessTd(IHtmlNode td, TableRow row)
		{
			if (td.HasChildren)
			{
				TableCell cell = new TableCell();
				Paragraph para = null;
				//Run run = null;
				
				foreach (IHtmlNode child in td.Children)
				{
					if (child.IsText)
					{
						if (para == null)
						{
							para = new Paragraph();
							ParagraphCreated(td, para);
							//para = CreateParagraph(td);
							//run = CreateRun(td, para);
						}
						
						Run run = para.AppendChild(new Run(new Text(child.InnerHtml)));
						RunCreated(child, run);
						//run.AppendChild(new Text(child.InnerHtml));
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
					if (string.Compare(td.Tag, "td", StringComparison.InvariantCultureIgnoreCase) == 0)
					{
						ProcessTd(td, row);
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
			return string.Compare(node.Tag, "table", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			if (node == null || parent == null || !CanConvert(node))
			{
				return;
			}
			
			//Parent.Current = null;
			paragraph = null;
			
			if (node.HasChildren)
			{
				Table table = new Table();
				
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

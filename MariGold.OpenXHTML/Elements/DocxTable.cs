namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxTable : WordElement
	{
		private void ProcessTd(HtmlNode td, TableRow row)
		{
			if (td.HasChildren)
			{
				TableCell cell = new TableCell();
				Run run = null;
				
				foreach (HtmlNode child in td.Children)
				{
					if (child.IsText)
					{
						if (run == null)
						{
							run = new Run();
						}
						
						run.AppendChild(new Text(child.InnerHtml));
					}
					else
					{
						if (run != null)
						{
							cell.Append(run);
						}
						
						ProcessChild(child, cell);
					}
				}
				
				if (run != null)
				{
					cell.Append(run);
				}
				
				row.Append(cell);
			}
		}
		
		private void ProcessTr(HtmlNode tr, Table table)
		{
			if (tr.HasChildren)
			{
				TableRow row = new TableRow();
				
				foreach (HtmlNode td in tr.Children)
				{
					if (string.Compare(td.Tag, "td", true) == 0)
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
		
		internal override bool IsBlockLine
		{
			get
			{
				return true;
			}
		}
		
		internal override bool CanConvert(HtmlNode node)
		{
			return string.Compare(node.Tag, "table", true) == 0;
		}
		
		internal override void Process(HtmlNode node, OpenXmlElement parent)
		{
			if (node == null || parent == null || string.Compare(node.Tag, "table", true) != 0)
			{
				return;
			}
			
			if (node.HasChildren)
			{
				Table table = parent.AppendChild(new Table());
				
				foreach (HtmlNode tr in node.Children)
				{
					if (string.Compare(tr.Tag, "tr", true) == 0)
					{
						ProcessTr(tr, table);
					}
				}
			}
		}
	}
}

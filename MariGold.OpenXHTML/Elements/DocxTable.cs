namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxTable : WordElement
	{
		private void ProcessTd(HtmlNode td)
		{
			if (td.HasChildren)
			{
				foreach (HtmlNode child in td.Children) 
				{
					
				}
			}
		}
		
		private void ProcessTr(HtmlNode tr)
		{
			if (tr.HasChildren)
			{
				TableRow row = new TableRow();
				
				foreach (HtmlNode td in tr.Children)
				{
					if (string.Compare(td.Tag, "td", true) == 0)
					{
						ProcessTd(td);
					}
				}
				
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
			
			Table table = parent.AppendChild(new Table());
			
			if (node.HasChildren)
			{
				foreach (HtmlNode tr in node.Children)
				{
					if (string.Compare(tr.Tag, "tr", true) == 0)
					{
						ProcessTr(tr);
					}
				}
			}
		}
	}
}

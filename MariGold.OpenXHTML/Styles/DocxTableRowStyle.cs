namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;
	using MariGold.HtmlParser;
	
	internal sealed class DocxTableRowStyle
	{
		internal void Process(TableRow row, DocxTableProperties docxProperties)
		{
			TableRowProperties trProperties = new TableRowProperties();
			
			if (docxProperties.CellSpacing != null)
			{
				trProperties.Append(new TableCellSpacing() {
					Width = DocxUnits.GetDxaFromPixel(docxProperties.CellSpacing.Value).ToString(),
					Type = TableWidthUnitValues.Dxa
				});
			}
			
			if (trProperties.ChildElements.Count > 0)
			{
				row.Append(trProperties);
			}
		}
	}
}

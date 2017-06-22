namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml.Wordprocessing;

	internal sealed class DocxTableCellStyle
	{
		private const string colspan = "colspan";

        private void ProcessBorders(DocxNode node, DocxTableProperties docxProperties,
			TableCellProperties cellProperties)
		{
            string borderStyle = node.ExtractStyleValue(DocxBorder.borderName);
            string leftBorder = node.ExtractStyleValue(DocxBorder.leftBorderName);
            string topBorder = node.ExtractStyleValue(DocxBorder.topBorderName);
            string rightBorder = node.ExtractStyleValue(DocxBorder.rightBorderName);
            string bottomBorder = node.ExtractStyleValue(DocxBorder.bottomBorderName);
			
			TableCellBorders cellBorders = new TableCellBorders();
			
			DocxBorder.ApplyBorders(cellBorders, borderStyle, leftBorder, topBorder, 
				rightBorder, bottomBorder, docxProperties.HasDefaultBorder);
			
			if (cellBorders.HasChildren)
			{
				cellProperties.Append(cellBorders);
			}
		}

        private void ProcessColSpan(DocxNode node, TableCellProperties cellProperties)
		{
			Int32 value;

            if (Int32.TryParse(node.ExtractAttributeValue(colspan), out value))
			{
				if (value > 1)
				{
					cellProperties.Append(new GridSpan() { Val = value });
				}
			}
		}

        private void ProcessWidth(DocxNode node, TableCellProperties cellProperties)
		{
            string width = node.ExtractStyleValue(DocxUnits.width);
			
			if (!string.IsNullOrEmpty(width))
			{
				decimal value;
				TableWidthUnitValues unit;
				
				if (DocxUnits.TableUnitsFromStyle(width, out value, out unit))
				{
					TableCellWidth cellWidth = new TableCellWidth() {
						Width = value.ToString(),
						Type = unit
					};
					
					cellProperties.Append(cellWidth);
				}
			}
		}

        private void ProcessVerticalAlignment(DocxNode node, TableCellProperties cellProperties)
		{
            string alignment = node.ExtractStyleValue(DocxAlignment.verticalAlign);
			
			if (!string.IsNullOrEmpty(alignment))
			{
				TableVerticalAlignmentValues value;
				
				if (DocxAlignment.GetCellVerticalAlignment(alignment, out value))
				{
					cellProperties.Append(new TableCellVerticalAlignment(){ Val = value });
				}
			}
		}
		
		internal bool HasRowSpan{ get; set; }
		
		internal void Process(TableCell cell, DocxTableProperties docxProperties, DocxNode node)
		{
			TableCellProperties cellProperties = new TableCellProperties();

            ProcessColSpan(node, cellProperties);
            ProcessWidth(node, cellProperties);
			
			if (HasRowSpan)
			{
				cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
			}
			
			//Processing border should be after colspan
            ProcessBorders(node, docxProperties, cellProperties);

            string backgroundColor = node.ExtractStyleValue(DocxColor.backGroundColor);
			
			if (!string.IsNullOrEmpty(backgroundColor))
			{
				DocxColor.ApplyBackGroundColor(backgroundColor, cellProperties);
			}

            ProcessVerticalAlignment(node, cellProperties);
			
			if (cellProperties.HasChildren)
			{
				cell.Append(cellProperties);
			}
		}

        internal static DocxNode GetHtmlNodeForTableCellContent(DocxNode node)
		{
            node.RemoveStyles(DocxBorder.borderName, DocxBorder.leftBorderName, DocxBorder.rightBorderName,
                DocxBorder.topBorderName, DocxBorder.bottomBorderName, DocxMargin.margin, DocxMargin.marginLeft,
                DocxMargin.marginRight, DocxMargin.marginTop, DocxMargin.marginBottom, DocxColor.backGroundColor);

            return node;
		}
	}
}

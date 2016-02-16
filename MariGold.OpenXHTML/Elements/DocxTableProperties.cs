namespace MariGold.OpenXHTML
{
	using System;
	
	internal sealed class DocxTableProperties
	{
		private bool hasDefaultHeader;
		private bool isCellHeader;
		private Int16? cellPadding;
		private Int16? cellSpacing;
		
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
	}
}

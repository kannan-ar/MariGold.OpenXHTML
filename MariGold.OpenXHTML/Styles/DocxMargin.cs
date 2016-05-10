namespace MariGold.OpenXHTML
{
	using System;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal sealed class DocxMargin
	{
		private readonly DocxNode node;
		
		internal const string margin = "margin";
		internal const string marginTop = "margin-top";
		internal const string marginBottom = "margin-bottom";
		internal const string marginLeft = "margin-left";
		internal const string marginRight = "margin-right";
		internal const string lineHeight = "line-height";
		
		internal DocxMargin(IHtmlNode node)
		{
			this.node = new DocxNode(node);
		}
		
		internal DocxMargin(DocxNode node)
		{
			this.node = node;
		}
		
		internal string GetTopMargin()
		{
			string topMargin = node.ExtractStyleValue(marginTop);
			
			if (string.IsNullOrEmpty(topMargin))
			{
				topMargin = node.ExtractStyleValue(margin);
			}
			
			return topMargin;
		}
		
		internal string GetBottomMargin()
		{
			string bottomMargin = node.ExtractStyleValue(marginBottom);
			
			if (string.IsNullOrEmpty(bottomMargin))
			{
				bottomMargin = node.ExtractStyleValue(margin);
			}
			
			return bottomMargin;
		}
		
		internal string GetLeftMargin()
		{
			string leftMargin = node.ExtractStyleValue(marginLeft);
			
			if (string.IsNullOrEmpty(leftMargin))
			{
				leftMargin = node.ExtractStyleValue(margin);
			}
			
			return leftMargin;
		}
		
		internal string GetRightMargin()
		{
			string rightMargin = node.ExtractStyleValue(marginRight);
			
			if (string.IsNullOrEmpty(rightMargin))
			{
				rightMargin = node.ExtractStyleValue(margin);
			}
			
			return rightMargin;
		}
		
		internal void SetLeftMargin(string value)
		{
			node.SetStyleValue(marginLeft, value);
		}
		
		internal void ProcessParagraphMargin(ParagraphProperties properties)
		{
			string topMargin = GetTopMargin();
			string bottomMargin = GetBottomMargin();
			string leftMargin = GetLeftMargin();
			string rightMargin = GetRightMargin();
			string line = node.ExtractStyleValue(lineHeight);
			
			if (!string.IsNullOrEmpty(topMargin) || !string.IsNullOrEmpty(bottomMargin) || !string.IsNullOrEmpty(line))
			{
				SpacingBetweenLines spacing = new SpacingBetweenLines();
				
				if (!string.IsNullOrEmpty(topMargin))
				{
                    int dxa = DocxUnits.GetDxaFromStyle(topMargin);

                    if (dxa != -1)
                    {
                        spacing.Before = dxa.ToString();
                    }
				}
				
				if (!string.IsNullOrEmpty(bottomMargin))
				{
                    int dxa = DocxUnits.GetDxaFromStyle(bottomMargin);

                    if (dxa != -1)
                    {
                        spacing.After = dxa.ToString();
                    }
				}
				
				if (!string.IsNullOrEmpty(line) && !line.CompareStringInvariantCultureIgnoreCase(DocxFont.normal))
				{
                    int dxa = DocxUnits.GetDxaFromStyle(line);

                    if (dxa != -1)
                    {
                        spacing.Line = dxa.ToString();
                    }
				}

                if (spacing.HasAttributes)
                {
                    properties.Append(spacing);
                }
			}
			
			if (!string.IsNullOrEmpty(leftMargin) || !string.IsNullOrEmpty(rightMargin))
			{
				Indentation ind = new Indentation();
				
				if (!string.IsNullOrEmpty(leftMargin))
				{
                    int dxa = DocxUnits.GetDxaFromStyle(leftMargin);

                    if (dxa != -1)
                    {
                        ind.Left = dxa.ToString();
                    }
				}
				
				if (!string.IsNullOrEmpty(rightMargin))
				{
                    int dxa = DocxUnits.GetDxaFromStyle(rightMargin);

                    if (dxa != -1)
                    {
                        ind.Right = dxa.ToString();
                    }
				}

                if (ind.HasAttributes)
                {
                    properties.Append(ind);
                }
			}
		}
		
		internal static void SetTopMargin(string style, ParagraphProperties properties)
		{
			Int32 dxa = DocxUnits.GetDxaFromStyle(style);
			
			if (dxa != -1)
			{
				SpacingBetweenLines spacing = new SpacingBetweenLines();
				
				spacing.Before = dxa.ToString();
				properties.Append(spacing);
			}
		}
		
		internal static void SetBottomMargin(string style, ParagraphProperties properties)
		{
			Int32 dxa = DocxUnits.GetDxaFromStyle(style);
			
			if (dxa != -1)
			{
				SpacingBetweenLines spacing = new SpacingBetweenLines();
				
				spacing.After = dxa.ToString();
				properties.Append(spacing);
			}
		}
	}
}

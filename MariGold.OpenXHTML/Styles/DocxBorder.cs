namespace MariGold.OpenXHTML
{
	using System;
	using System.Collections;
	using System.Text.RegularExpressions;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.Collections.Generic;
	
	internal static class DocxBorder
	{
		private static IDictionary<string,BorderValues> boderStyles;
		
		static DocxBorder()
		{
			SetBorderValues();
		}
		
		private static void SetBorderValues()
		{
			boderStyles = new Dictionary<string,BorderValues>();
			
			boderStyles.Add("dotted", BorderValues.Dotted);
			boderStyles.Add("dashed", BorderValues.Dashed);
			boderStyles.Add("solid", BorderValues.Single);
			boderStyles.Add("double", BorderValues.Double);
			boderStyles.Add("groove", BorderValues.ThreeDEngrave);
			boderStyles.Add("ridge", BorderValues.ThreeDEmboss);
			boderStyles.Add("inset", BorderValues.Inset);
			boderStyles.Add("outset", BorderValues.Outset);
			boderStyles.Add("none", BorderValues.None);
			boderStyles.Add("hidden", BorderValues.None);
			
		}
		
		private static string GetBorderWidth(ref string borderStyle)
		{
			string width = string.Empty;
			
			Match match = Regex.Match(borderStyle, "\\d+((px)|(pt)|(cm)|(em))", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
			
			if (match.Success)
			{
				Match intValue = Regex.Match(match.Value, "\\d+");
				
				if (intValue.Success)
				{
					width = intValue.Value;
				}
				
				borderStyle = borderStyle.Replace(match.Value, string.Empty);
			}
			
			return width;
		}
		
		private static BorderValues GetBorderStyle(ref string borderStyle)
		{
			foreach (var style in boderStyles)
			{
				int index = borderStyle.IndexOf(style.Key, StringComparison.InvariantCultureIgnoreCase);
					
				if (index != -1)
				{
					borderStyle = borderStyle.Replace(style.Key, string.Empty);
						
					return style.Value;
				}
			}
			
			return BorderValues.None;
		}
		
		private static void GetBorderProperties(string borderStyle, out BorderValues borderType, out string color, out UInt32 width)
		{
			string _width = GetBorderWidth(ref borderStyle);
			borderType = GetBorderStyle(ref borderStyle);
			
			UInt32.TryParse(_width, out width);
			color = DocxColor.GetHexColor(borderStyle.Trim().ToLower());
		}
		
		private static T GetBorderType<T>(BorderValues borderType, string color, UInt32 width)
			where T : BorderType, new()
		{
			if (borderType == BorderValues.None)
			{
				return null;
			}
			
			if (string.IsNullOrEmpty(color))
			{
				return null;
			}
			
			if (width == 0)
			{
				return null;
			}
			
			var border = new T(){ Val = borderType, Color = color, Size = width, Space = (UInt32Value)0U };
			
			return border;
		}
		
		internal static void ApplyDefaultBorders(OpenXmlCompositeElement element)
		{
			TopBorder topBorder = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
			LeftBorder leftBorder = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
			BottomBorder bottomBorder = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
			RightBorder rightBorder = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
				
			element.Append(topBorder);
			element.Append(leftBorder);
			element.Append(bottomBorder);
			element.Append(rightBorder);
		}
		
		internal static void ApplyBorders(OpenXmlCompositeElement element, 
			string boderStyle, 
			string leftBorderStyle, 
			string topBorderStyle, 
			string rightBorderStyle, 
			string bottomBorderStyle)
		{
			BorderValues borderValue;
			string color;
			UInt32 width;
			
			GetBorderProperties(boderStyle, out borderValue, out color, out width);
			
			if (!string.IsNullOrEmpty(topBorderStyle))
			{
				BorderValues topBorderValue;
				string topColor;
				UInt32 topWidth;
			
				GetBorderProperties(topBorderStyle, out topBorderValue, out topColor, out topWidth);
				
				TopBorder topBorder = GetBorderType<TopBorder>(topBorderValue, topColor, topWidth);
				
				if (topBorder != null)
				{
					element.Append(topBorder);
				}
			}
			else
			{
				TopBorder topBorder = GetBorderType<TopBorder>(borderValue, color, width);
				
				if (topBorder != null)
				{
					element.Append(topBorder);
				}
			}
			
			if (!string.IsNullOrEmpty(leftBorderStyle))
			{
				BorderValues leftBorderValue;
				string leftColor;
				UInt32 leftWidth;
			
				GetBorderProperties(leftBorderStyle, out leftBorderValue, out leftColor, out leftWidth);
				
				LeftBorder leftBorder = GetBorderType<LeftBorder>(leftBorderValue, leftColor, leftWidth);
				
				if (leftBorder != null)
				{
					element.Append(leftBorder);
				}
			}
			else
			{
				LeftBorder leftBorder = GetBorderType<LeftBorder>(borderValue, color, width);
				
				if (leftBorder != null)
				{
					element.Append(leftBorder);
				}
			}
			
			if (!string.IsNullOrEmpty(bottomBorderStyle))
			{
				BorderValues bottomBorderValue;
				string bottomColor;
				UInt32 bottomWidth;
			
				GetBorderProperties(bottomBorderStyle, out bottomBorderValue, out bottomColor, out bottomWidth);
				
				BottomBorder bottomBorder = GetBorderType<BottomBorder>(bottomBorderValue, bottomColor, bottomWidth);
				
				if (bottomBorder != null)
				{
					element.Append(bottomBorder);
				}
			}
			else
			{
				BottomBorder bottomBorder = GetBorderType<BottomBorder>(borderValue, color, width);
				
				if (bottomBorder != null)
				{
					element.Append(bottomBorder);
				}
			}
			
			if (!string.IsNullOrEmpty(rightBorderStyle))
			{
				BorderValues rightBorderValue;
				string rightColor;
				UInt32 rightWidth;
			
				GetBorderProperties(rightBorderStyle, out rightBorderValue, out rightColor, out rightWidth);
				
				RightBorder rightBorder = GetBorderType<RightBorder>(rightBorderValue, rightColor, rightWidth);
				
				if (rightBorder != null)
				{
					element.Append(rightBorder);
				}
			}
			else
			{
				RightBorder rightBorder = GetBorderType<RightBorder>(borderValue, color, width);
				
				if (rightBorder != null)
				{
					element.Append(rightBorder);
				}
			}
		}
	}
}

namespace MariGold.OpenXHTML
{
	using System;
	using System.Text.RegularExpressions;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Wordprocessing;
	using System.Collections.Generic;
	
	internal static class DocxBorder
	{
		private static IDictionary<string,BorderValues> boderStyles;
		
		internal const string borderName = "border";
		internal const string leftBorderName = "border-left";
		internal const string topBorderName = "border-top";
		internal const string rightBorderName = "border-right";
		internal const string bottomBorderName = "border-bottom";
		
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
		
		private static void ApplyBorder<T>(string cssStyle, OpenXmlCompositeElement element)
			where T : BorderType, new()
		{
			BorderValues borderValue;
			string color;
			UInt32 width;
			
			GetBorderProperties(cssStyle, out borderValue, out color, out width);
				
			T border = GetBorderType<T>(borderValue, color, width);
				
			if (border != null)
			{
				element.Append(border);
			}
		}
		
		private static void ApplyDefaultBorder<T>(OpenXmlCompositeElement element)
			where T : BorderType, new()
		{
			T border = new T() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
			element.Append(border);
		}
		
		internal static void ApplyDefaultBorders(OpenXmlCompositeElement element)
		{
			ApplyDefaultBorder<TopBorder>(element);
			ApplyDefaultBorder<LeftBorder>(element);
			ApplyDefaultBorder<BottomBorder>(element);
			ApplyDefaultBorder<RightBorder>(element);
		}
		
		internal static void ApplyBorders(OpenXmlCompositeElement element, 
			string borderStyle, 
			string leftBorderStyle, 
			string topBorderStyle, 
			string rightBorderStyle, 
			string bottomBorderStyle,
			bool useDefaultBorder)
		{
			BorderValues borderValue = BorderValues.None;
			string color = string.Empty;
			UInt32 width = 0;
			bool hasBorder = false;
			
			if (!string.IsNullOrEmpty(borderStyle))
			{
				hasBorder = true;
				GetBorderProperties(borderStyle, out borderValue, out color, out width);
			}
			
			if (!string.IsNullOrEmpty(topBorderStyle))
			{
				ApplyBorder<TopBorder>(topBorderStyle, element);
			}
			else if (hasBorder)
			{
				TopBorder topBorder = GetBorderType<TopBorder>(borderValue, color, width);
				
				if (topBorder != null)
				{
					element.Append(topBorder);
				}
			}
			else if(useDefaultBorder)
			{
				ApplyDefaultBorder<TopBorder>(element);
			}
			
			if (!string.IsNullOrEmpty(leftBorderStyle))
			{
				ApplyBorder<LeftBorder>(leftBorderStyle, element);
			}
			else if (hasBorder)
			{
				LeftBorder leftBorder = GetBorderType<LeftBorder>(borderValue, color, width);
				
				if (leftBorder != null)
				{
					element.Append(leftBorder);
				}
			}
			else if(useDefaultBorder)
			{
				ApplyDefaultBorder<LeftBorder>(element);
			}
			
			if (!string.IsNullOrEmpty(bottomBorderStyle))
			{
				ApplyBorder<BottomBorder>(bottomBorderStyle, element);
			}
			else if (hasBorder)
			{
				BottomBorder bottomBorder = GetBorderType<BottomBorder>(borderValue, color, width);
				
				if (bottomBorder != null)
				{
					element.Append(bottomBorder);
				}
			}
			else if(useDefaultBorder)
			{
				ApplyDefaultBorder<BottomBorder>(element);
			}
			
			if (!string.IsNullOrEmpty(rightBorderStyle))
			{
				ApplyBorder<RightBorder>(rightBorderStyle, element);
			}
			else if (hasBorder)
			{
				RightBorder rightBorder = GetBorderType<RightBorder>(borderValue, color, width);
				
				if (rightBorder != null)
				{
					element.Append(rightBorder);
				}
			}
			else if(useDefaultBorder)
			{
				ApplyDefaultBorder<RightBorder>(element);
			}
		}
	}
}

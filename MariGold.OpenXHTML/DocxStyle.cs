namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	using System.Collections.Generic;
	
	internal abstract class DocxStyle<T>
		where T : OpenXmlElement
	{
		protected const string backGroundColor = "background-color";
		protected const string color = "color";
		protected const string fontFamily = "font-family";
		
		internal abstract void Process(T element, Dictionary<string,string> styles);
	}
}

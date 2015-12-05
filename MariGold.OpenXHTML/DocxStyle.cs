namespace MariGold.OpenXHTML
{
	using System;
	using DocumentFormat.OpenXml;
	using System.Collections.Generic;
	
	internal abstract class DocxStyle<T>
		where T : OpenXmlElement
	{
		internal abstract void Process(T element, Dictionary<string,string> styles);
	}
}

namespace MariGold.OpenXHTML
{
	using System;
	
	/// <summary>
	/// 
	/// </summary>
	public sealed class WordDocument
	{
		private readonly IHtmlParser parser;
		
		public WordDocument(IHtmlParser parser)
		{
			this.parser = parser;
		}
		
		public void Save(string path)
		{
			
		}
	}
}

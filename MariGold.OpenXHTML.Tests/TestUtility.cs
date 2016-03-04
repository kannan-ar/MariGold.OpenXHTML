namespace MariGold.OpenXHTML.Tests
{
	using System;
	using NUnit.Framework;
	using Word = DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class TestUtility
	{
		internal static void TestBorder<T>(T border, Word.BorderValues val, string color, UInt32 width)
			where T : Word.BorderType
		{
			Assert.AreEqual(border.Val.Value, val);
			Assert.AreEqual(border.Color.Value, color);
			Assert.AreEqual(border.Size, width);
		}
		
		internal static void TestTableCell(this Word.TableCell cell, Int32 childCount, string content)
		{
			Assert.IsNotNull(cell);
			Assert.AreEqual(childCount, cell.ChildElements.Count);
				
			Word.Paragraph para = cell.ChildElements[0] as Word.Paragraph;
				
			Assert.IsNotNull(para);
			Assert.AreEqual(1, para.ChildElements.Count);
				
			Word.Run run = para.ChildElements[0] as Word.Run;
				
			Assert.IsNotNull(run);
			Assert.AreEqual(1, run.ChildElements.Count);
				
			Word.Text text = run.ChildElements[0] as Word.Text;
				
			Assert.IsNotNull(text);
			Assert.AreEqual(0, text.ChildElements.Count);
			Assert.AreEqual(content, text.InnerText);
		}
	}
}

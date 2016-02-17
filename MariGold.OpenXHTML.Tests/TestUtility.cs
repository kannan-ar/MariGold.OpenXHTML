namespace MariGold.OpenXHTML.Tests
{
	using System;
	using NUnit.Framework;
	using DocumentFormat.OpenXml.Wordprocessing;
	
	internal static class TestUtility
	{
		internal static void TestBorder<T>(T border, BorderValues val, string color, UInt32 width)
			where T : BorderType
		{
			Assert.AreEqual(border.Val.Value, val);
			Assert.AreEqual(border.Color.Value, color);
			Assert.AreEqual(border.Size, width);
		}
	}
}

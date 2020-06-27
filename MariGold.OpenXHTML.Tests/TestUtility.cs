namespace MariGold.OpenXHTML.Tests
{
	using System;
	using NUnit.Framework;
	using Word = DocumentFormat.OpenXml.Wordprocessing;
	using System.Collections.Generic;
	using DocumentFormat.OpenXml.Validation;
	using System.Linq;
	using System.Text;
	using System.IO;
	
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
		
		internal static void PrintValidationErrors(this IEnumerable<ValidationErrorInfo> errors)
		{
			if (errors.Any())
			{
				StringBuilder sb = new StringBuilder();
				
				foreach (var error in errors)
				{
					sb.AppendLine(error.Description);
				}
				
				Assert.Fail(sb.ToString());
			}
		}
		
        internal static string GetPath(string folderPath)
        {
            string path = Environment.CurrentDirectory;
			int index = path.IndexOf("bin");
			path = path.Remove(index);
            return path + folderPath;
        }

		internal static string GetHtmlFromFile(string folderPath)
		{
			string html = string.Empty;
			string path = Environment.CurrentDirectory;
			int index = path.IndexOf("bin");

			path = path.Remove(index);
			
			using (StreamReader sr = new StreamReader(path + folderPath))
			{
				html = sr.ReadToEnd();
			}
			
			return html;
		}
	}
}

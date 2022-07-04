namespace MariGold.OpenXHTML.Tests
{
    using DocumentFormat.OpenXml.Validation;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using Xunit;
    using Word = DocumentFormat.OpenXml.Wordprocessing;

    internal static class TestUtility
    {
        internal static void Equal(int expected, uint actual)
        {
            Assert.Equal(expected, (int)actual);
        }

        internal static void TestBorder<T>(T border, Word.BorderValues val, string color, UInt32 width)
            where T : Word.BorderType
        {
            Assert.Equal(border.Val.Value, val);
            Assert.Equal(border.Color.Value, color);
            Assert.Equal(border.Size.Value, width);
        }

        internal static void TestTableCell(this Word.TableCell cell, Int32 childCount, string content)
        {
            Assert.NotNull(cell);
            Assert.Equal(childCount, cell.ChildElements.Count);

            Word.Paragraph para = cell.ChildElements[0] as Word.Paragraph;

            Assert.NotNull(para);
            Assert.Equal(1, para.ChildElements.Count);

            Word.Run run = para.ChildElements[0] as Word.Run;

            Assert.NotNull(run);
            Assert.Equal(1, run.ChildElements.Count);

            Word.Text text = run.ChildElements[0] as Word.Text;

            Assert.NotNull(text);
            Assert.Equal(0, text.ChildElements.Count);
            Assert.Equal(content, text.InnerText);
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

                Assert.True(false, sb.ToString());
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

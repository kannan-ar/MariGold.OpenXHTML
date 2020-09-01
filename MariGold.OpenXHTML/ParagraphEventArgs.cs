namespace MariGold.OpenXHTML
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;

    public class ParagraphEventArgs : EventArgs
    {
        public Paragraph Paragraph { get; }

        public ParagraphEventArgs(Paragraph paragraph)
        {
            this.Paragraph = paragraph;
        }
    }
}

namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;

    public class ParagraphEventArgs : EventArgs
    {
        public Paragraph Paragraph { get; }

        public ParagraphEventArgs(Paragraph paragraph)
        {
            this.Paragraph = paragraph;
        }
    }
}

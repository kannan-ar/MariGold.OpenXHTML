namespace MariGold.OpenXHTML
{
    using System;
    using DocumentFormat.OpenXml.Wordprocessing;

    public class ParagraphEventArgs : EventArgs
    {
        private readonly Paragraph paragraph;

        public Paragraph Paragraph
        {
            get
            {
                return paragraph;
            }
        }

        public ParagraphEventArgs(Paragraph paragraph)
        {
            this.paragraph = paragraph;
        }
    }
}

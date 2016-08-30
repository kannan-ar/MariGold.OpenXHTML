namespace MariGold.OpenXHTML
{
    using System;
    using System.Web;
    using System.Text.RegularExpressions;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal abstract class DocxElement
    {
        protected readonly IOpenXmlContext context;
        internal EventHandler<ParagraphEventArgs> ParagraphCreated;

        protected void RunCreated(DocxNode node, Run run)
        {
            DocxRunStyle style = new DocxRunStyle();
            style.Process(run, node);
        }

        protected void OnParagraphCreated(DocxNode node, Paragraph para)
        {
            DocxParagraphStyle style = new DocxParagraphStyle();
            style.Process(para, node);

            if (ParagraphCreated != null)
            {
                ParagraphCreated(this, new ParagraphEventArgs(para));
            }
        }

        protected void ProcessChild(DocxNode node, ref Paragraph paragraph)
        {
            DocxElement element = context.Convert(node);

            if (element != null)
            {
                if (ParagraphCreated != null)
                {
                    element.ParagraphCreated = ParagraphCreated;
                }

                element.Process(node, ref paragraph);
            }
        }

        protected void ProcessTextElement(DocxNode node)
        {
            ITextElement element = context.ConvertTextElement(node);

            if (element != null)
            {
                element.Process(node);
            }
        }

        protected void ProcessTextChild(DocxNode node)
        {
            foreach (DocxNode child in node.Children)
            {
                if (child.IsText && !IsEmptyText(child.InnerHtml))
                {
                    Run run = node.Parent.AppendChild(new Run(new Text()
                    {
                        Text = ClearHtml(child.InnerHtml),
                        Space = SpaceProcessingModeValues.Preserve
                    }));

                    RunCreated(node, run);
                }
                else
                {
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessTextElement(child);
                }
            }
        }

        protected string CleanUrl(string url)
        {
            if (url.StartsWith("//") && !string.IsNullOrEmpty(context.UriSchema))
            {
                url = string.Concat(context.UriSchema, ":" + url);
            }

            if (Uri.IsWellFormedUriString(url, UriKind.Relative) && !string.IsNullOrEmpty(context.BaseURL))
            {
                url = string.Concat(context.BaseURL, url);
            }

            return url;
        }

        protected bool IsHidden(DocxNode node)
        {
            if (node == null)
            {
                return false;
            }

            string display = node.ExtractStyleValue("display");
            return display.CompareStringInvariantCultureIgnoreCase("none");
        }

        protected void ProcessElement(DocxNode node, ref Paragraph paragraph)
        {
            foreach (DocxNode child in node.Children)
            {
                if (child.IsText)
                {
                    ProcessParagraph(child, node, ref paragraph);
                }
                else
                {
                    child.ParagraphNode = node.ParagraphNode;
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessChild(child, ref paragraph);
                }
            }
        }

        protected void ProcessParagraph(DocxNode child, DocxNode node, ref Paragraph paragraph)
        {
            if (!IsEmptyText(child.InnerHtml))
            {
                if (paragraph == null)
                {
                    paragraph = node.Parent.AppendChild(new Paragraph());
                    OnParagraphCreated(node.ParagraphNode, paragraph);
                }

                Run run = paragraph.AppendChild(new Run(new Text()
                {
                    Text = ClearHtml(child.InnerHtml),
                    Space = SpaceProcessingModeValues.Preserve
                }));

                RunCreated(node, run);
            }
        }

        internal DocxElement(IOpenXmlContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            this.context = context;
        }

        internal string ClearHtml(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return string.Empty;
            }

            html = HttpUtility.HtmlDecode(html);
            html = html.Replace("&nbsp;", " ");
            html = html.Replace("&amp;", "&");

            Regex regex = new Regex(Environment.NewLine + "\\s+");
            Match match = regex.Match(html);

            while (match.Success)
            {
                //match.Length - 1 for leave a single space. Otherwise the sentences will collide.
                html = html.Remove(match.Index, match.Length - 1);
                match = regex.Match(html);
            }

            html = html.Replace(Environment.NewLine, string.Empty);

            return html;
        }

        internal bool IsEmptyText(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return true;
            }

            html = html.Replace(Environment.NewLine, string.Empty);

            if (string.IsNullOrEmpty(html.Trim()))
            {
                return true;
            }

            return false;
        }

        internal abstract bool CanConvert(DocxNode node);

        internal abstract void Process(DocxNode node, ref Paragraph paragraph);
    }
}

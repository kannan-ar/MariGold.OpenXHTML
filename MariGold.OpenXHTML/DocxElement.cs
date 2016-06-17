namespace MariGold.OpenXHTML
{
    using System;
    using System.Text.RegularExpressions;
    using MariGold.HtmlParser;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;

    internal abstract class DocxElement
    {
        protected readonly IOpenXmlContext context;

        protected void RunCreated(IHtmlNode node, Run run)
        {
            DocxRunStyle style = new DocxRunStyle();
            style.Process(run, node);
        }

        protected void ParagraphCreated(IHtmlNode node, Paragraph para)
        {
            DocxParagraphStyle style = new DocxParagraphStyle();
            style.Process(para, node);
        }

        protected void ProcessChild(DocxProperties properties, ref Paragraph paragraph)
        {
            DocxElement element = context.Convert(properties.CurrentNode);

            if (element != null)
            {
                element.Process(properties, ref paragraph);
            }
        }

        protected void ProcessTextElement(DocxProperties properties)
        {
            ITextElement element = context.ConvertTextElement(properties.CurrentNode);

            if (element != null)
            {
                element.Process(properties);
            }
        }

        protected void ProcessTextChild(DocxProperties properties)
        {
            foreach (IHtmlNode child in properties.CurrentNode.Children)
            {
                if (child.IsText && !IsEmptyText(child.InnerHtml))
                {
                    Run run = properties.Parent.AppendChild(new Run(new Text()
                    {
                        Text = ClearHtml(child.InnerHtml),
                        Space = SpaceProcessingModeValues.Preserve
                    }));

                    RunCreated(properties.CurrentNode, run);
                }
                else
                {
                    ProcessTextElement(new DocxProperties(child, properties.Parent));
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

        protected bool IsHidden(IHtmlNode node)
        {
            if (node == null)
            {
                return false;
            }

            DocxNode docxNode = new DocxNode(node);
            string display = docxNode.ExtractStyleValue("display");
            return display.CompareStringInvariantCultureIgnoreCase("none");
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

        internal abstract bool CanConvert(IHtmlNode node);

        internal abstract void Process(DocxProperties properties, ref Paragraph paragraph);
    }
}

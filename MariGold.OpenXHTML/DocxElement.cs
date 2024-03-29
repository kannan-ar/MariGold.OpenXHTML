﻿namespace MariGold.OpenXHTML
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using MariGold.OpenXHTML.Styles;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Text.RegularExpressions;

    internal abstract class DocxElement
    {
        protected const string whiteSpace = " ";
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

            ParagraphCreated?.Invoke(this, new ParagraphEventArgs(para));
        }

        protected void ProcessChild(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            DocxElement element = context.Convert(node);

            if (element != null)
            {
                if (ParagraphCreated != null)
                {
                    element.ParagraphCreated = ParagraphCreated;
                }

                element.Process(node, ref paragraph, properties);
            }
        }

        protected void ProcessTextElement(DocxNode node, Dictionary<string, object> properties)
        {
            ITextElement element = context.ConvertTextElement(node);

            if (element != null)
            {
                element.Process(node, properties);
            }
        }

        protected void ProcessTextChild(DocxNode node, Dictionary<string, object> properties)
        {
            foreach (DocxNode child in node.Children)
            {
                if (child.IsText && !IsEmptyText(child.InnerHtml))
                {
                    Run run = node.Parent.AppendChild(new Run(new[] {new Text()
                    {
                        Text = ClearHtml(child.InnerHtml),
                        Space = SpaceProcessingModeValues.Preserve
                    } }));

                    RunCreated(node, run);
                }
                else
                {
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessTextElement(child, properties);
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
            return display.CompareStringOrdinalIgnoreCase("none");
        }

        protected void ProcessElement(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            foreach (DocxNode child in node.Children)
            {
                if (child.IsText)
                {
                    ProcessParagraph(child, node, node.ParagraphNode, ref paragraph);
                }
                else
                {
                    child.ParagraphNode = node.ParagraphNode;
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    ProcessChild(child, ref paragraph, properties);
                }
            }
        }

        protected void ProcessBlockElement(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            foreach (DocxNode child in node.Children)
            {
                if (child.IsText)
                {
                    ProcessParagraph(child, node, node, ref paragraph);
                }
                else
                {
                    //ProcessChild forwards the incomming parent to the child element. So any div element inside this div
                    //creates a new paragraph on the parent element.
                    child.ParagraphNode = node;
                    child.Parent = node.Parent;
                    node.CopyExtentedStyles(child);
                    node.ApplyBlockStyles(child);

                    ProcessChild(child, ref paragraph, properties);
                }
            }
        }

        protected void ProcessParagraph(DocxNode child, DocxNode node, DocxNode paragraphNode, ref Paragraph paragraph)
        {
            if (!IsEmptyText(child, out string text))
            {
                if (paragraph == null)
                {
                    paragraph = node.Parent.AppendChild(new Paragraph());
                    OnParagraphCreated(paragraphNode, paragraph);
                }

                Run run = paragraph.AppendChild(new Run(new[] { new Text()
                {
                    Text = ClearHtml(text),
                    Space = SpaceProcessingModeValues.Preserve
                }}));

                RunCreated(node, run);
            }
        }

        protected bool TryCreateAbsoluteUri(string relativeUrl, out Uri uri)
        {
            if (relativeUrl.StartsWith("//") && !string.IsNullOrEmpty(context.UriSchema))
            {
                relativeUrl = string.Concat(context.UriSchema, ":" + relativeUrl);
            }

            if (Uri.IsWellFormedUriString(relativeUrl, UriKind.Relative) && !string.IsNullOrEmpty(context.ImagePath))
            {
                relativeUrl = string.Concat(context.ImagePath,
                    (!context.ImagePath.EndsWith("/") && !relativeUrl.StartsWith("/")) ? "/" : string.Empty,
                    relativeUrl);
            }
            else if (Uri.IsWellFormedUriString(relativeUrl, UriKind.Relative) && !string.IsNullOrEmpty(context.BaseURL))
            {
                relativeUrl = string.Concat(context.BaseURL,
                    (!context.BaseURL.EndsWith("/") && !relativeUrl.StartsWith("/")) ? "/" : string.Empty,
                    relativeUrl);
            }

            return Uri.TryCreate(WebUtility.UrlDecode(relativeUrl), UriKind.Absolute, out uri);
        }

        protected bool TryCreateFromEncodedString(string data, out string value)
        {
            Match match = Regex.Match(data, "data(\\s*):");
            value = string.Empty;

            if (match.Success)
            {
                value = data[(match.Index + match.Length)..];
            }

            return match.Success;
        }

        protected Stream GetStream(Uri uri)
        {
            using WebClient client = new WebClient() { Encoding = System.Text.Encoding.UTF8 };
            client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:31.0) Gecko/20100101 Firefox/31.0");
            client.UseDefaultCredentials = true;

            return client.OpenRead(uri);
        }

        private protected DocxElement(IOpenXmlContext context)
        {
            this.context = context ?? throw new ArgumentNullException("context");
        }

        internal string ClearHtml(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return string.Empty;
            }

            html = WebUtility.HtmlDecode(html);
            html = html.Replace("&nbsp;", whiteSpace);
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

        internal bool IsEmptyText(DocxNode node, out string text)
        {
            text = string.Empty;

            if (string.IsNullOrEmpty(node.InnerHtml))
            {
                return true;
            }

            text = node.InnerHtml.Replace(Environment.NewLine, string.Empty);

            if (!string.IsNullOrEmpty(text.Trim()))
            {
                return false;
            }
            else if (!string.IsNullOrEmpty(text) &&
                node.Previous != null && !node.Previous.IsText && !node.Previous.InnerHtml.EndsWith(whiteSpace) &&
                node.Next != null && !node.Next.IsText && !node.Next.InnerHtml.StartsWith(whiteSpace))
            {
                text = whiteSpace;
                return false;
            }

            return true;
        }

        internal abstract bool CanConvert(DocxNode node);

        internal abstract void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties);
    }
}

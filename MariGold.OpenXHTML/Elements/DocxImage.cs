namespace MariGold.OpenXHTML
{
    using System;
    using System.Net;
    using System.IO;
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using A = DocumentFormat.OpenXml.Drawing;
    using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
    using System.Drawing;
    using System.Text.RegularExpressions;
    using System.Collections.Generic;

    internal sealed class DocxImage : DocxElement, ITextElement
    {
        private ImagePartType GetImagePartType(string src)
        {
            string ext = Path.GetExtension(src);

            if (ext != null)
            {
                ext = ext.ToLower().Replace(".", string.Empty);
            }

            Enum.TryParse(ext, true, out ImagePartType type);

            return type;
        }

        private Drawing CreateDrawingFromStream(string src, ImagePartType imagePartType, Func<Stream> getStream)
        {
            long cx;
            long cy;

            using (Stream stream = getStream())
            {
                using (Bitmap bitmap = new Bitmap(stream))
                {
                    cx = (long)bitmap.Width * (long)((float)914400 / bitmap.HorizontalResolution);
                    cy = (long)bitmap.Height * (long)((float)914400 / bitmap.VerticalResolution);
                }
            }

            using (Stream stream = getStream())
            {
                ImagePart imagePart = context.MainDocumentPart.AddImagePart(imagePartType);

                imagePart.FeedData(stream);

                var image = new Drawing(
                                new DW.Inline(
                                    new DW.Extent() { Cx = cx, Cy = cy },
                                    new DW.EffectExtent()
                                    {
                                        LeftEdge = 0L,
                                        TopEdge = 0L,
                                        RightEdge = 0L,
                                        BottomEdge = 0L
                                    },
                                    new DW.DocProperties()
                                    {
                                        Id = (UInt32Value)1U,
                                        Name = "Picture 1"
                                    },
                                    new DW.NonVisualGraphicFrameDrawingProperties(
                                        new A.GraphicFrameLocks() { NoChangeAspect = true }),
                                    new A.Graphic(
                                        new A.GraphicData(
                                            new PIC.Picture(
                                                new PIC.NonVisualPictureProperties(
                                                    new PIC.NonVisualDrawingProperties()
                                                    {
                                                        Id = (UInt32Value)0U,
                                                        Name = Path.GetFileName(src)
                                                    },
                                                    new PIC.NonVisualPictureDrawingProperties()),
                                                    new PIC.BlipFill(
                                                    new A.Blip(
                                                        new A.BlipExtensionList(
                                                            new A.BlipExtension()
                                                            {
                                                                Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                            })
                                                    )
                                                    {
                                                        Embed = context.MainDocumentPart.GetIdOfPart(imagePart),
                                                        CompressionState = A.BlipCompressionValues.Print
                                                    },
                                                    new A.Stretch(
                                                        new A.FillRectangle())),
                                                    new PIC.ShapeProperties(
                                                    new A.Transform2D(
                                                        new A.Offset() { X = 0L, Y = 0L },
                                                        new A.Extents() { Cx = cx, Cy = cy }),
                                                    new A.PresetGeometry(
                                                        new A.AdjustValueList()
                                                    )
                                                    { Preset = A.ShapeTypeValues.Rectangle }))
                                        )
                                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                                )
                                {
                                    DistanceFromTop = (UInt32Value)0U,
                                    DistanceFromBottom = (UInt32Value)0U,
                                    DistanceFromLeft = (UInt32Value)0U,
                                    DistanceFromRight = (UInt32Value)0U
                                });

                return image;
            }
        }

        private Drawing CreateDrawingFromData(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            value = value.Trim();

            Match match = Regex.Match(value, "image/([a-zA-Z]+)");
            string ext = match.Groups.Count > 1 ? match.Groups[1].Value : string.Empty;
            int dataIndex = value.LastIndexOf(',');

            if (string.IsNullOrEmpty(ext) || dataIndex == -1 || dataIndex + 1 >= value.Length)
            {
                return null;
            }

            Enum.TryParse(ext, true, out ImagePartType type);

            return CreateDrawingFromStream(value, type, () =>
            {
                var bytes = Convert.FromBase64String(value.Substring(dataIndex + 1).Trim());
                return new MemoryStream(bytes);
            });
        }

        private Drawing CreateDrawingFromAbsoluteUri(string src, Uri uri)
        {
            return CreateDrawingFromStream(src, GetImagePartType(src), () =>
            {
                return GetStream(uri);
            });
        }

        private Drawing PrepareImage(string src)
        {
            if (TryCreateFromEncodedString(src, out string value))
            {
                return CreateDrawingFromData(value);
            }
            else if (TryCreateAbsoluteUri(WebUtility.UrlEncode(src), out Uri uri))
            {
                return CreateDrawingFromAbsoluteUri(src, uri);
            }

            return null;
        }

        internal DocxImage(IOpenXmlContext context)
            : base(context)
        {
        }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Equals(node.Tag, "img", StringComparison.OrdinalIgnoreCase);
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph, Dictionary<string, object> properties)
        {
            if (IsHidden(node))
            {
                return;
            }

            string src = node.ExtractAttributeValue("src");

            if (!string.IsNullOrEmpty(src))
            {
                try
                {
                    Drawing drawing = PrepareImage(src);
                    
                    if (drawing != null)
                    {
                        if (paragraph == null)
                        {
                            paragraph = node.Parent.AppendChild(new Paragraph());
                            OnParagraphCreated(node, paragraph);
                        }

                        Run run = paragraph.AppendChild(new Run(drawing));
                        RunCreated(node, run);
                    }
                }
                catch
                {
                    return;//fails silently?
                }
            }
        }

        bool ITextElement.CanConvert(DocxNode node)
        {
            return CanConvert(node);
        }

        void ITextElement.Process(DocxNode node, Dictionary<string, object> properties)
        {
            if (IsHidden(node))
            {
                return;
            }

            string src = node.ExtractAttributeValue("src");

            if (!string.IsNullOrEmpty(src))
            {
                try
                {
                    Drawing drawing = PrepareImage(src);

                    if (drawing != null)
                    {
                        Run run = node.Parent.AppendChild(new Run(drawing));
                        RunCreated(node, run);
                    }
                }
                catch
                {
                    return;//fails silently?
                }
            }
        }
    }
}

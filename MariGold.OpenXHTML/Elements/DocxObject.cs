namespace MariGold.OpenXHTML
{
    using System;
    using System.IO;
    using System.Drawing;
    using System.Threading.Tasks;

    using DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Packaging;
    using V = DocumentFormat.OpenXml.Vml;
    using OVML = DocumentFormat.OpenXml.Vml.Office;

    internal sealed class DocxObject : DocxElement, ITextElement
    {
        private Func<IOpenXmlContext, string> getRelationshipId = (context) => string.Concat("rId", (++context.RelationshipId).ToString());

        private bool IsImage(string filePath)
        {
            return filePath.HasStringContains(".jpg", ".bmp", ".gif", ".png", ".tiff");
        }

        private bool TryFormat(string filePath, out string contentType, out string progId)
        {
            contentType = string.Empty;
            progId = string.Empty;
            
            if (filePath.HasStringContains(".docx"))
            {
                contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                progId = "Word.Document.12";
                return true;
            }
            else if (filePath.HasStringContains(".xlsx"))
            {
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                progId = "Excel.Sheet.12";
                return true;
            }
            else if (filePath.HasStringContains(".pptx"))
            {
                contentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
                progId = "PowerPoint.Show.12";
                return true;
            }

            return false;
        }

        private MemoryStream GetImageStream(string fileName, out int width, out int height)
        {
            width = 0;
            height = 0;

            using (Image img = new Bitmap(1, 1))
            {
                using (Graphics drawing = Graphics.FromImage(img))
                {
                    SizeF textSize = drawing.MeasureString(fileName, SystemFonts.DefaultFont);

                    width = (int)textSize.Width;
                    height = (int)textSize.Height;
                }
            }

            MemoryStream mem = new MemoryStream();

            using (Image img = new Bitmap(width, height))
            {
                using (Graphics drawing = Graphics.FromImage(img))
                {
                    drawing.Clear(System.Drawing.Color.White);

                    Brush textBrush = new SolidBrush(System.Drawing.Color.Blue);
                    drawing.DrawString(fileName, SystemFonts.DefaultFont, textBrush, 0, 0);
                    img.Save(mem, System.Drawing.Imaging.ImageFormat.Jpeg);

                }
            }

            mem.Position = 0;

            return mem;
        }

        private EmbeddedObject ProcessObject(string filePath, string contentType, string progId)
        {

            if (TryCreateAbsoluteUri(filePath, out Uri uri))
            {
                int width, height;
                string relationshipId = string.Concat("rId", (++context.RelationshipId).ToString());

                EmbeddedPackagePart embeddedPackagePart = context.MainDocumentPart.AddNewPart<EmbeddedPackagePart>(contentType,
                    relationshipId);

                using (Stream stream = GetStream(uri))
                {
                    embeddedPackagePart.FeedData(stream);
                }

                ImagePart imagePart = context.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);

                using (MemoryStream mem = GetImageStream(Path.GetFileName(filePath), out width, out height))
                {
                    imagePart.FeedData(mem);
                }

                V.Shape shape = new V.Shape() { Id = string.Concat("_x0000_i", context.RelationshipId.ToString()), Ole = false, Type = "#_x0000_t75", Style = "width:" + width.ToString() + "pt;height:" + height.ToString() + "pt" };
                V.ImageData imageData = new V.ImageData() { Title = "", RelationshipId = context.MainDocumentPart.GetIdOfPart(imagePart) };

                shape.Append(imageData);

                EmbeddedObject embeddedObject = new EmbeddedObject() { DxaOriginal = "9360", DyaOriginal = "450" };
                embeddedObject.Append(shape);

                OVML.OleObject oleObject = new OVML.OleObject() { Type = OVML.OleValues.Embed, ProgId = progId, ShapeId = string.Concat("_x0000_i", context.RelationshipId.ToString()), DrawAspect = OVML.OleDrawAspectValues.Content, ObjectId = string.Concat("_", context.RelationshipId.ToString()), Id = relationshipId };

                embeddedObject.Append(oleObject);

                return embeddedObject;
            }

            return null;
        }

        internal DocxObject(IOpenXmlContext context) : base(context) { }

        internal override bool CanConvert(DocxNode node)
        {
            return string.Equals(node.Tag, "object", StringComparison.OrdinalIgnoreCase);
        }

        internal override void Process(DocxNode node, ref Paragraph paragraph)
        {
            if (IsHidden(node))
            {
                return;
            }

            string contentType;
            string progId;
            string data = node.ExtractAttributeValue("data");

            if (IsImage(data))
            {
                var interchanger = context.GetInterchanger();
                interchanger.ProcessImage(context, data, node, ref paragraph);
            }
            else if (TryFormat(data, out contentType, out progId))
            {
                EmbeddedObject embeddedObject = ProcessObject(data, contentType, progId);

                if (embeddedObject != null)
                {
                    if (paragraph == null)
                    {
                        paragraph = node.Parent.AppendChild(new Paragraph());
                        OnParagraphCreated(node, paragraph);
                    }

                    Run run = paragraph.AppendChild(new Run(embeddedObject));
                    RunCreated(node, run);
                }
            }
        }

        bool ITextElement.CanConvert(DocxNode node)
        {
            return string.Equals(node.Tag, "object", StringComparison.OrdinalIgnoreCase);
        }

        void ITextElement.Process(DocxNode node)
        {
            if (IsHidden(node))
            {
                return;
            }

            string contentType;
            string progId;
            string data = node.ExtractAttributeValue("data");

            if (IsImage(data))
            {
                var interchanger = context.GetInterchanger();
                interchanger.ProcessImage(context, data, node);
            }
            else if (TryFormat(data, out contentType, out progId))
            {
                EmbeddedObject embeddedObject = ProcessObject(data, contentType, progId);

                if (embeddedObject != null)
                {
                    Run run = node.Parent.AppendChild(new Run(embeddedObject));
                    RunCreated(node, run);
                }
            }
        }
    }
}

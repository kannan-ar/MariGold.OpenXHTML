namespace MariGold.OpenXHTML
{
	using System;
	using System.Net;
	using System.IO;
	using MariGold.HtmlParser;
	using DocumentFormat.OpenXml;
	using DocumentFormat.OpenXml.Packaging;
	using DocumentFormat.OpenXml.Wordprocessing;
	using A = DocumentFormat.OpenXml.Drawing;
	using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
	using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
	using System.Drawing;

	internal sealed class DocxImage : DocxElement
	{
		private ImagePartType GetImagePartType(string src)
		{
			ImagePartType type;

			string ext = Path.GetExtension(src);

			if (ext != null)
			{
				ext = ext.ToLower().Replace(".", string.Empty);
			}

			Enum.TryParse<ImagePartType>(ext, out type);
			
			return type;
		}
		
		private Drawing CreateDrawingFromAbsoluteUri(string src)
		{
			long cx;
			long cy;
			
			WebClient client = new WebClient() { Encoding = System.Text.Encoding.UTF8 };
			client.Headers.Add("user-agent", "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:31.0) Gecko/20100101 Firefox/31.0");
			client.UseDefaultCredentials = true;
			
			using (Stream stream = client.OpenRead(new Uri(src)))
			{
				using (Bitmap bitmap = new Bitmap(stream))
				{
					cx = (long)bitmap.Width * (long)((float)914400 / bitmap.HorizontalResolution);
					cy = (long)bitmap.Height * (long)((float)914400 / bitmap.VerticalResolution);
				}
			}
				
			using (Stream stream = client.OpenRead(new Uri(src)))
			{
				
				ImagePart imagePart = context.MainDocumentPart.AddImagePart(GetImagePartType(src));
					
				imagePart.FeedData(stream);
					
				var image = new Drawing(
					            new DW.Inline(
						            new DW.Extent() { Cx = cx, Cy = cy },
						            new DW.EffectExtent() {
							LeftEdge = 0L,
							TopEdge = 0L,
							RightEdge = 0L,
							BottomEdge = 0L
						},
						            new DW.DocProperties() {
							Id = (UInt32Value)1U,
							Name = "Picture 1"
						},
						            new DW.NonVisualGraphicFrameDrawingProperties(
							            new A.GraphicFrameLocks() { NoChangeAspect = true }),
						            new A.Graphic(
							            new A.GraphicData(
								            new PIC.Picture(
									            new PIC.NonVisualPictureProperties(
										            new PIC.NonVisualDrawingProperties() {
											Id = (UInt32Value)0U,
											Name = Path.GetFileName(src)
										},
										            new PIC.NonVisualPictureDrawingProperties()),
									            new PIC.BlipFill(
										            new A.Blip(
											            new A.BlipExtensionList(
												            new A.BlipExtension() {
													Uri =
                                                                            "{28A0092B-C50C-407E-A947-70E740481C1C}"
												})
										            ) {
											Embed = context.MainDocumentPart.GetIdOfPart(imagePart),
											CompressionState =
                                                                    A.BlipCompressionValues.Print
										},
										            new A.Stretch(
											            new A.FillRectangle())),
									            new PIC.ShapeProperties(
										            new A.Transform2D(
											            new A.Offset() { X = 0L, Y = 0L },
											            new A.Extents() { Cx = cx, Cy = cy }),
										            new A.PresetGeometry(
											            new A.AdjustValueList()
										            ) { Preset = A.ShapeTypeValues.Rectangle }))
							            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
					            ) {
						DistanceFromTop = (UInt32Value)0U,
						DistanceFromBottom = (UInt32Value)0U,
						DistanceFromLeft = (UInt32Value)0U,
						DistanceFromRight = (UInt32Value)0U,
						EditId = "50D07946"
					});
					
				return image;
			}
		}
		
		private Drawing PrepareImage(string src)
		{
			if (Uri.IsWellFormedUriString(src, UriKind.Relative) && !string.IsNullOrEmpty(context.ImagePath))
			{
				src = string.Concat(context.ImagePath, src);
			}
			
			if (Uri.IsWellFormedUriString(src, UriKind.Absolute))
			{
				return CreateDrawingFromAbsoluteUri(src);
			}
			
			return null;
		}
		
		internal DocxImage(IOpenXmlContext context)
			: base(context)
		{
		}
		
		internal override bool CanConvert(IHtmlNode node)
		{
			return string.Compare(node.Tag, "img", StringComparison.InvariantCultureIgnoreCase) == 0;
		}
		
		internal override void Process(IHtmlNode node, OpenXmlElement parent, ref Paragraph paragraph)
		{
			DocxNode docxNode = new DocxNode(node);
			
			string src = docxNode.ExtractAttributeValue("src");
			
			if (!string.IsNullOrEmpty(src))
			{
				Drawing drawing = PrepareImage(src);
				
				if (drawing != null)
				{
					if (paragraph == null)
					{
						paragraph = parent.AppendChild(new Paragraph());
						ParagraphCreated(node, paragraph);
					}
					
					Run run = paragraph.AppendChild(new Run(drawing));
					RunCreated(node, run);
				}
			}
		}
	}
}

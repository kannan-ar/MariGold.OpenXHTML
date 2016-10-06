##MariGold.OpenXHTML
OpenXHTML is a wrapper library for Open XML SDK to convert HTML documents into Open XML word documents. It simply encapsulated the complexity of Open XML yet exposes its properties for manipulation.

###Installing via NuGet

In Package Manager Console, enter the following command:
```
Install-Package MariGold.OpenXHTML
```
###Usage
To create an Open XML word document using the OpenXHTML, use the following code.

```csharp
using MariGold.OpenXHTML;

WordDocument doc = new WordDocument("sample.docx");
doc.Save();
```
To parse the HTML and convert into an Open XML document, use the following code.

```csharp
using MariGold.OpenXHTML;

WordDocument doc = new WordDocument("sample.docx");
doc.Process(new HtmlParser("<div>sample text</div>"));
doc.Save();
```
Once the HTML is processed, you can access the Open XML document using the following properties in `WordDocument`.

```csharp
public WordprocessingDocument WordprocessingDocument { get; }
public MainDocumentPart MainDocumentPart { get; }
public Document Document { get; }
```
For example, if you want to append a paragraph at the document body, try the following code.
```csharp
using MariGold.OpenXHTML;
using DocumentFormat.OpenXml.Wordprocessing;

WordDocument doc = new WordDocument("sample.docx");
doc.Process(new HtmlParser("<div>sample text</div>"));
doc.Document.Body.AppendChild<Paragraph>(new Paragraph(new Run(new Text("added text"))));
doc.Save();
```
If you want to create an Open XML document in memory, use the following code.

```csharp
using (MemoryStream mem = new MemoryStream())
{
	WordDocument doc = new WordDocument(mem);
	doc.Save();
}
			
```
####Relative Images
Sometimes the given html document may contains relative image url. OpenXHTML cannot process the images in such cases. An image path can be set to avoid this issue.

```csharp
using MariGold.OpenXHTML;

WordDocument doc = new WordDocument("sample.docx");
doc.ImagePath = "http:\\abc.com";
doc.Process(new HtmlParser("<img src="sample.png" />"));
doc.Save();
```

You can also assign any physical address on image path.
```csharp
doc.ImagePath = @"C:\Img";
```

####Base URL
Like image path, an html document may also contain links with relative path. This can resolve using the `BaseURL` property.

```csharp
using MariGold.OpenXHTML;

WordDocument doc = new WordDocument("sample.docx");
doc.BaseURL = "http:\\abc.com";
doc.Process(new HtmlParser("<a href="index.htm">sample</a>"));
doc.Save();
```
Also, if there any relative images in the given html document and `ImagePath` is not assigned, OpenXHTML will attempt to use `BaseURL` to solve relative image paths.

####HTML Parsing
OpenXHTML is using a built-in HTML and CSS parser (MariGold.HtmlParser) which can complectly replace with any external HTML and CSS parser. The `Process` method in `WordDocument` expects an `IParser` interface type implementation to process the HTML and CSS.
```csharp
public void Process(IParser parser);
```

```csharp
interface IParser
{
	string BaseURL { get; set; }
	string UriSchema { get; set; }

	decimal CalculateRelativeChildFontSize(string parentFontSize, string childFontSize);
	IHtmlNode FindBodyOrFirstElement();
}
```
Here is the structure of `IParser`. The `BaseURL` and `UriSchema` are two simple properties to store the base url address and uri schema to process the html images and links. The `CalculateRelativeChildFontSize` is used to calculate the relative child font size. For example, in the below html, the font size for the h1 tag will be 20 pixel. 

```html
<div style="font-size:16px"><h1>sample</h1></div>
```

If you don't want to reimplement this functionality, you can simple use the CSSUtility class in MariGold.HtmlParser

```csharp
using MariGold.HtmlParser;

CSSUtility.CalculateRelativeChildFontSize(parentFontSize, childFontSize);
```

The `FindBodyOrFirstElement` method expected to return an IHtmlNode implementation of html body and the hierarchy of its child elements. If the document does not have body element, then return the first root element.

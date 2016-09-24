##MariGold.OpenXHTML
OpenXHTML is a wrapper library for Open XML SDK to convert HTML documents into Open XML word documents. It simply encapsulated the complexity of Open XML yet exposes the Open XML document objects for manipulation.

###Installing via NuGet

In Package Manager Console, enter the following command:
```
Install-Package MariGold.OpenXHTML
```

###Usage
To create an Open XML word document using the OpenXHTML, use the following code.

```csharp
WordDocument doc = new WordDocument("sample.docx");
doc.Save();
```

To parse the HTML and convert into an Open XML document, use the following code.

```csharp
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

####Base URL

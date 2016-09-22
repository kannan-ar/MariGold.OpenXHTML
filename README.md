##MariGold.OpenXHTML
OpenXHTML is a wrapper class of Open XML SDK to convert HTML documents to Open XML word documents. 

###Installing via NuGet

In Package Manager Console, enter the following command:
```
Install-Package MariGold.OpenXHTML
```

###Usage
To create an Open XML word document using the OpexXHTML, use the following code.

```csharp
WordDocument doc = new WordDocument("sample.docx");
doc.Save();
```

###Relative Images

###Base URL

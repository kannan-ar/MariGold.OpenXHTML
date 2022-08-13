namespace MariGold.OpenXHTML
{
    using MariGold.HtmlParser;

    public interface IParser
    {
        string BaseURL { get; set; }
        string UriSchema { get; set; }
        IHtmlNode FindBodyOrFirstElement();
        decimal CalculateRelativeChildFontSize(string parentFontSize, string childFontSize);
    }
}

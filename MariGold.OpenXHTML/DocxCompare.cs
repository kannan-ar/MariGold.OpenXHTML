namespace MariGold.OpenXHTML
{
    using System;

    internal static class DocxCompare
    {
        internal static bool CompareStringInvariantCultureIgnoreCase(this string source, string value)
        {
            if(string.IsNullOrEmpty(source) || string.IsNullOrEmpty(value))
            {
                return false;
            }

            return string.Compare(source, value, StringComparison.InvariantCultureIgnoreCase) == 0;
        }

        internal static bool IsImage(this string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }

            string[] imageExtensions = new string[] { ".jpg", ".bmp", ".gif", ".png" };

            foreach(string ext in imageExtensions)
            {
                if(path.Contains(ext))
                {
                    return true;
                }
            }

            return false;
        }
    }
}

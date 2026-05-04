using System.Text;

namespace DotnetPoi.SS.Util;

public static class LocaleUtil
{
    // RegisterProvider must run before GetEncoding(1252) because:
    // - Static field initializers run before the static constructor in C#.
    // - On non-Windows platforms, code-page encodings are not available by default.
    // Solution: initialize Charset1252 inside the static constructor, after registration.
    static LocaleUtil()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Charset1252 = Encoding.GetEncoding(1252);
    }

    public static readonly Encoding Charset1252;
}

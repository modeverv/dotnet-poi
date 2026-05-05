using System.Text;
using DotnetPoi.SS.Xml;
using Xunit;

namespace DotnetPoi.SS.Tests.Xml;

public class PoiXmlWriterDeclarationProfileTests
{
    [Fact]
    public void WriteStartDocument_XmlBeansSpreadsheetPart_WritesUtf8DeclarationNewlineAndNoStandalone()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument(PoiXmlDeclarationProfile.XmlBeansSpreadsheetPart);
            writer.WriteStartElement("worksheet");
            writer.WriteEndElement();
        }

        Assert.Equal("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<worksheet/>", builder.ToString());
    }

    [Fact]
    public void WriteStartDocument_OpcPackagePart_WritesUtf8StandaloneDeclarationWithoutRootNewline()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument(PoiXmlDeclarationProfile.OpcPackagePart);
            writer.WriteStartElement("Types");
            writer.WriteEndElement();
        }

        Assert.Equal("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types/>", builder.ToString());
    }

    [Fact]
    public void WriteStartDocument_ProfileAfterDeclaration_Throws()
    {
        var builder = new StringBuilder();
        using var writer = new PoiXmlWriter(new StringWriter(builder));

        writer.WriteStartDocument(PoiXmlDeclarationProfile.OpcPackagePart);

        Assert.Throws<InvalidOperationException>(() =>
            writer.WriteStartDocument(PoiXmlDeclarationProfile.XmlBeansSpreadsheetPart));
    }
}

using System.Text;
using DotnetPoi.SS.Xml;
using Xunit;

namespace DotnetPoi.SS.Tests.Xml;

public class PoiXmlWriterFactoryTests
{
    [Fact]
    public void Create_XmlBeansSpreadsheetPart_WritesProfileDeclarationBeforeRoot()
    {
        var builder = new StringBuilder();
        using (var writer = PoiXmlWriterFactory.Create(
            new StringWriter(builder),
            PoiXmlDeclarationProfile.XmlBeansSpreadsheetPart))
        {
            writer.WriteStartElement("worksheet");
            writer.WriteEndElement();
        }

        Assert.Equal("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<worksheet/>", builder.ToString());
    }

    [Fact]
    public void CreateForOoxmlPackagePart_OpcParts_UseStandaloneProfile()
    {
        Assert.Equal(PoiXmlDeclarationProfile.OpcPackagePart,
            PoiXmlWriterFactory.GetDeclarationProfileForOoxmlPackagePart("[Content_Types].xml"));
        Assert.Equal(PoiXmlDeclarationProfile.OpcPackagePart,
            PoiXmlWriterFactory.GetDeclarationProfileForOoxmlPackagePart("_rels/.rels"));
        Assert.Equal(PoiXmlDeclarationProfile.OpcPackagePart,
            PoiXmlWriterFactory.GetDeclarationProfileForOoxmlPackagePart("xl/_rels/workbook.xml.rels"));
        Assert.Equal(PoiXmlDeclarationProfile.OpcPackagePart,
            PoiXmlWriterFactory.GetDeclarationProfileForOoxmlPackagePart("docProps/core.xml"));
    }

    [Fact]
    public void CreateForOoxmlPackagePart_XmlBeansParts_UseXmlBeansProfile()
    {
        Assert.Equal(PoiXmlDeclarationProfile.XmlBeansSpreadsheetPart,
            PoiXmlWriterFactory.GetDeclarationProfileForOoxmlPackagePart("xl/workbook.xml"));
        Assert.Equal(PoiXmlDeclarationProfile.XmlBeansSpreadsheetPart,
            PoiXmlWriterFactory.GetDeclarationProfileForOoxmlPackagePart("docProps/app.xml"));
        Assert.Equal(PoiXmlDeclarationProfile.XmlBeansSpreadsheetPart,
            PoiXmlWriterFactory.GetDeclarationProfileForOoxmlPackagePart("/word/document.xml"));
        Assert.Equal(PoiXmlDeclarationProfile.XmlBeansSpreadsheetPart,
            PoiXmlWriterFactory.GetDeclarationProfileForOoxmlPackagePart("ppt/presentation.xml"));
    }

    [Fact]
    public void CreateForOoxmlPackagePart_PackagePart_WritesProfileDeclarationBeforeRoot()
    {
        var builder = new StringBuilder();
        using (var writer = PoiXmlWriterFactory.CreateForOoxmlPackagePart(
            new StringWriter(builder),
            "_rels/.rels"))
        {
            writer.WriteStartElement("Relationships");
            writer.WriteEndElement();
        }

        Assert.Equal("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships/>",
            builder.ToString());
    }
}

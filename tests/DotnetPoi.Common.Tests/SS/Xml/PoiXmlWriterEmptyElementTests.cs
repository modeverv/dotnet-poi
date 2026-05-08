using System.Text;
using DotnetPoi.SS.Xml;
using Xunit;

namespace DotnetPoi.SS.Tests.Xml;

public class PoiXmlWriterEmptyElementTests
{
    [Fact]
    public void WriteEndElement_RootEmptyElement_WritesSlashWithoutLeadingSpace()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartElement("root");
            writer.WriteEndElement();
        }

        Assert.Equal("<root/>", builder.ToString());
    }

    [Fact]
    public void WriteEndElement_NestedEmptyElement_WritesSlashWithoutLeadingSpace()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartElement("root");
            writer.WriteStartElement("child");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        Assert.Equal("<root><child/></root>", builder.ToString());
    }

    [Fact]
    public void WriteEndElement_PrefixedEmptyElement_WritesSlashWithoutLeadingSpace()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartElement("x", "node");
            writer.WriteEndElement();
        }

        Assert.Equal("<x:node/>", builder.ToString());
    }

    [Fact]
    public void WriteEndElement_AttributedEmptyElement_WritesSlashWithoutLeadingSpace()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartElement("node");
            writer.WriteAttributeString("id", "rId1");
            writer.WriteEndElement();
        }

        Assert.Equal("<node id=\"rId1\"/>", builder.ToString());
    }

    [Fact]
    public void WriteEndElement_MultipleAttributesEmptyElement_WritesSlashWithoutLeadingSpace()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartElement("Relationship");
            writer.WriteAttributeString("Id", "rId1");
            writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            writer.WriteAttributeString("Target", "worksheets/sheet1.xml");
            writer.WriteEndElement();
        }

        Assert.Equal(
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>",
            builder.ToString());
    }

    [Fact]
    public void WriteEndElement_PrefixedAttributedEmptyElement_WritesSlashWithoutLeadingSpace()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartElement("a", "picLocks");
            writer.WriteAttributeString("noChangeAspect", "1");
            writer.WriteEndElement();
        }

        Assert.Equal("<a:picLocks noChangeAspect=\"1\"/>", builder.ToString());
    }

    [Fact]
    public void WriteEndElement_AfterEmptyStringWrite_WritesSlashWithoutLeadingSpace()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartElement("root");
            writer.WriteString(string.Empty);
            writer.WriteEndElement();
        }

        Assert.Equal("<root/>", builder.ToString());
    }
}

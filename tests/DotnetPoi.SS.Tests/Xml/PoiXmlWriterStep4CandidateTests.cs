using System.Text;
using DotnetPoi.SS.Xml;
using Xunit;

namespace DotnetPoi.SS.Tests.Xml;

public class PoiXmlWriterStep4CandidateTests
{
    [Fact]
    public void Write_EmptyStringText_UsesEmptyElementFixtureStyle()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartElement("t");
            writer.WriteString(string.Empty);
            writer.WriteEndElement();
        }

        Assert.Equal("<t/>", builder.ToString());
    }

    [Fact]
    public void Write_ContentTypes_MultiSheet_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: true);
            writer.WriteStartElement("Types");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");

            WriteDefault(writer,
                contentType: "application/vnd.openxmlformats-package.relationships+xml",
                extension: "rels");
            WriteDefault(writer, contentType: "application/xml", extension: "xml");

            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                partName: "/docProps/app.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-package.core-properties+xml",
                partName: "/docProps/core.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                partName: "/xl/sharedStrings.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                partName: "/xl/styles.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                partName: "/xl/workbook.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                partName: "/xl/worksheets/sheet1.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                partName: "/xl/worksheets/sheet2.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                partName: "/xl/worksheets/sheet3.xml");

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("multi-sheet__[Content_Types].xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_ContentTypes_Namespaces_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: true);
            writer.WriteStartElement("Types");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");

            WriteDefault(writer, contentType: "image/png", extension: "png");
            WriteDefault(writer,
                contentType: "application/vnd.openxmlformats-package.relationships+xml",
                extension: "rels");
            WriteDefault(writer,
                contentType: "application/vnd.openxmlformats-officedocument.vmlDrawing",
                extension: "vml");
            WriteDefault(writer, contentType: "application/xml", extension: "xml");

            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                partName: "/docProps/app.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-package.core-properties+xml",
                partName: "/docProps/core.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
                partName: "/xl/comments1.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.drawing+xml",
                partName: "/xl/drawings/drawing1.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                partName: "/xl/sharedStrings.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                partName: "/xl/styles.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                partName: "/xl/workbook.xml");
            WriteOverride(writer,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                partName: "/xl/worksheets/sheet1.xml");

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("namespaces__[Content_Types].xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_WorkbookRels_MultiSheet_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: true);
            writer.WriteStartElement("Relationships");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

            WriteRelationship(writer, "rId1", "sharedStrings.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
            WriteRelationship(writer, "rId2", "styles.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
            WriteRelationship(writer, "rId3", "worksheets/sheet1.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            WriteRelationship(writer, "rId4", "worksheets/sheet2.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            WriteRelationship(writer, "rId5", "worksheets/sheet3.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("multi-sheet__xl___rels__workbook.xml.rels"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_Workbook_MultiSheet_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: false);
            writer.WriteString("\n");
            writer.WriteStartElement("workbook");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            writer.WriteStartElement("workbookPr");
            writer.WriteAttributeString("date1904", "false");
            writer.WriteEndElement();

            writer.WriteStartElement("bookViews");
            writer.WriteStartElement("workbookView");
            writer.WriteAttributeString("activeTab", "0");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("sheets");
            WriteSheet(writer, "Alpha", "rId3", "1");
            WriteSheet(writer, "Beta", "rId4", "2");
            WriteSheet(writer, "Gamma", "rId5", "3");
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("multi-sheet__xl__workbook.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_Worksheet_InlineStrings_PreservesBodyNewlines()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: false);
            writer.WriteString("\n");
            writer.WriteStartElement("worksheet");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            writer.WriteStartElement("dimension");
            writer.WriteAttributeString("ref", "A1:B1");
            writer.WriteEndElement();

            writer.WriteStartElement("sheetViews");
            writer.WriteStartElement("sheetView");
            writer.WriteAttributeString("workbookViewId", "0");
            writer.WriteAttributeString("tabSelected", "true");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("sheetFormatPr");
            writer.WriteAttributeString("defaultRowHeight", "15.0");
            writer.WriteEndElement();

            writer.WriteStartElement("sheetData");
            writer.WriteString("\n");
            writer.WriteStartElement("row");
            writer.WriteAttributeString("r", "1");
            writer.WriteString("\n");
            WriteInlineCell(writer, "A1", "Inline String");
            WriteInlineCell(writer, "B1", "More Inline");
            writer.WriteEndElement();
            writer.WriteString("\n");
            writer.WriteEndElement();

            WritePageMargins(writer);
            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("inline-strings__xl__worksheets__sheet1.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_Worksheet_Namespaces_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: false);
            writer.WriteString("\n");
            writer.WriteStartElement("worksheet");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            writer.WriteStartElement("dimension");
            writer.WriteAttributeString("ref", "A1");
            writer.WriteEndElement();

            writer.WriteStartElement("sheetViews");
            writer.WriteStartElement("sheetView");
            writer.WriteAttributeString("workbookViewId", "0");
            writer.WriteAttributeString("tabSelected", "true");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("sheetFormatPr");
            writer.WriteAttributeString("defaultRowHeight", "15.0");
            writer.WriteEndElement();

            writer.WriteStartElement("sheetData");
            writer.WriteStartElement("row");
            writer.WriteAttributeString("r", "1");
            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", "A1");
            writer.WriteAttributeString("t", "s");
            writer.WriteAttributeString("s", "0");
            writer.WriteStartElement("v");
            writer.WriteString("0");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            WritePageMargins(writer);

            writer.WriteStartElement("drawing");
            writer.WriteAttributeString("r", "id", "rId1");
            writer.WriteEndElement();

            writer.WriteStartElement("legacyDrawing");
            writer.WriteAttributeString("r", "id", "rId3");
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("namespaces__xl__worksheets__sheet1.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_Drawing_Namespaces_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: false);
            writer.WriteString("\n");
            writer.WriteStartElement("xdr", "wsDr");
            writer.WriteAttributeString("xmlns:xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            writer.WriteAttributeString("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            writer.WriteAttributeString("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            writer.WriteStartElement("xdr", "twoCellAnchor");
            writer.WriteAttributeString("editAs", "twoCell");

            writer.WriteStartElement("xdr", "from");
            WriteTextElement(writer, "xdr", "col", "0");
            WriteTextElement(writer, "xdr", "colOff", "0");
            WriteTextElement(writer, "xdr", "row", "0");
            WriteTextElement(writer, "xdr", "rowOff", "0");
            writer.WriteEndElement();

            writer.WriteStartElement("xdr", "to");
            WriteTextElement(writer, "xdr", "col", "0");
            WriteTextElement(writer, "xdr", "colOff", "0");
            WriteTextElement(writer, "xdr", "row", "0");
            WriteTextElement(writer, "xdr", "rowOff", "0");
            writer.WriteEndElement();

            writer.WriteStartElement("xdr", "pic");
            writer.WriteStartElement("xdr", "nvPicPr");
            writer.WriteStartElement("xdr", "cNvPr");
            writer.WriteAttributeString("id", "1");
            writer.WriteAttributeString("name", "Picture 1");
            writer.WriteAttributeString("descr", "Picture");
            writer.WriteEndElement();
            writer.WriteStartElement("xdr", "cNvPicPr");
            writer.WriteStartElement("a", "picLocks");
            writer.WriteAttributeString("noChangeAspect", "true");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("xdr", "blipFill");
            writer.WriteStartElement("a", "blip");
            writer.WriteAttributeString("r", "embed", "rId1");
            writer.WriteEndElement();
            writer.WriteStartElement("a", "stretch");
            writer.WriteStartElement("a", "fillRect");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("xdr", "spPr");
            writer.WriteStartElement("a", "xfrm");
            writer.WriteStartElement("a", "off");
            writer.WriteAttributeString("x", "0");
            writer.WriteAttributeString("y", "0");
            writer.WriteEndElement();
            writer.WriteStartElement("a", "ext");
            writer.WriteAttributeString("cx", "0");
            writer.WriteAttributeString("cy", "0");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteStartElement("a", "prstGeom");
            writer.WriteAttributeString("prst", "rect");
            writer.WriteStartElement("a", "avLst");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndElement();

            writer.WriteStartElement("xdr", "clientData");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("namespaces__xl__drawings__drawing1.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_Comments_Namespaces_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: false);
            writer.WriteString("\n");
            writer.WriteStartElement("comments");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            writer.WriteStartElement("authors");
            writer.WriteStartElement("author");
            writer.WriteString(string.Empty);
            writer.WriteEndElement();
            WriteTextElement(writer, null, "author", "poi");
            writer.WriteEndElement();

            writer.WriteStartElement("commentList");
            writer.WriteStartElement("comment");
            writer.WriteAttributeString("ref", "A1");
            writer.WriteAttributeString("authorId", "1");
            writer.WriteStartElement("text");
            WriteTextElement(writer, null, "t", "Comment text");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("namespaces__xl__comments1.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    private static void WriteDefault(PoiXmlWriter writer, string contentType, string extension)
    {
        writer.WriteStartElement("Default");
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteAttributeString("Extension", extension);
        writer.WriteEndElement();
    }

    private static void WriteOverride(PoiXmlWriter writer, string contentType, string partName)
    {
        writer.WriteStartElement("Override");
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteAttributeString("PartName", partName);
        writer.WriteEndElement();
    }

    private static void WriteRelationship(PoiXmlWriter writer, string id, string target, string type)
    {
        writer.WriteStartElement("Relationship");
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Target", target);
        writer.WriteAttributeString("Type", type);
        writer.WriteEndElement();
    }

    private static void WriteSheet(PoiXmlWriter writer, string name, string relationshipId, string sheetId)
    {
        writer.WriteStartElement("sheet");
        writer.WriteAttributeString("name", name);
        writer.WriteAttributeString("r", "id", relationshipId);
        writer.WriteAttributeString("sheetId", sheetId);
        writer.WriteEndElement();
    }

    private static void WriteInlineCell(PoiXmlWriter writer, string address, string text)
    {
        writer.WriteStartElement("c");
        writer.WriteAttributeString("r", address);
        writer.WriteAttributeString("t", "inlineStr");
        writer.WriteStartElement("is");
        WriteTextElement(writer, null, "t", text);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WritePageMargins(PoiXmlWriter writer)
    {
        writer.WriteStartElement("pageMargins");
        writer.WriteAttributeString("bottom", "0.75");
        writer.WriteAttributeString("footer", "0.3");
        writer.WriteAttributeString("header", "0.3");
        writer.WriteAttributeString("left", "0.7");
        writer.WriteAttributeString("right", "0.7");
        writer.WriteAttributeString("top", "0.75");
        writer.WriteEndElement();
    }

    private static void WriteTextElement(PoiXmlWriter writer, string? prefix, string localName, string text)
    {
        if (prefix == null)
        {
            writer.WriteStartElement(localName);
        }
        else
        {
            writer.WriteStartElement(prefix, localName);
        }

        writer.WriteString(text);
        writer.WriteEndElement();
    }
}

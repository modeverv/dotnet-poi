using System.Text;
using DotnetPoi.SS.Xml;
using Xunit;

namespace DotnetPoi.SS.Tests.Xml;

public class PoiXmlWriterCellTypesTests
{
    [Fact]
    public void Write_Rels_CellTypes_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: true);
            writer.WriteStartElement("Relationships");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

            WriteRelationship(writer, "rId1", "xl/workbook.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            WriteRelationship(writer, "rId2", "docProps/app.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
            WriteRelationship(writer, "rId3", "docProps/core.xml",
                "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("cell-types___rels__.rels"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_DocPropsApp_CellTypes_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: false);
            writer.WriteString("\n");
            writer.WriteStartElement("Properties");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");

            writer.WriteStartElement("Application");
            writer.WriteString("Apache POI");
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("cell-types__docProps__app.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_DocPropsCore_CellTypes_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: true);
            writer.WriteStartElement("cp", "coreProperties");
            writer.WriteAttributeString("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
            writer.WriteAttributeString("xmlns:dc", "http://purl.org/dc/elements/1.1/");
            writer.WriteAttributeString("xmlns:dcterms", "http://purl.org/dc/terms/");
            writer.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");

            var fixturePath = XmlFixturePaths.GetFixturePath("cell-types__docProps__core.xml");
            var timestamp = PoiXmlWriterCellZeroTests.ExtractTimestampFromCoreXmlFixture(fixturePath);

            writer.WriteStartElement("dcterms", "created");
            writer.WriteAttributeString("xsi", "type", "dcterms:W3CDTF");
            writer.WriteString(timestamp);
            writer.WriteEndElement();

            writer.WriteStartElement("dc", "creator");
            writer.WriteString("Apache POI");
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("cell-types__docProps__core.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_WorkbookRels_CellTypes_MatchesFixture()
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

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("cell-types__xl___rels__workbook.xml.rels"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_SharedStrings_CellTypes_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: false);
            writer.WriteString("\n");
            writer.WriteStartElement("sst");
            writer.WriteAttributeString("count", "1");
            writer.WriteAttributeString("uniqueCount", "1");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            writer.WriteStartElement("si");
            writer.WriteStartElement("t");
            writer.WriteString("Text");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("cell-types__xl__sharedStrings.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_Styles_CellTypes_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: false);
            writer.WriteString("\n");
            writer.WriteStartElement("styleSheet");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            writer.WriteStartElement("numFmts");
            writer.WriteAttributeString("count", "1");
            writer.WriteStartElement("numFmt");
            writer.WriteAttributeString("numFmtId", "164");
            writer.WriteAttributeString("formatCode", "yyyy-mm-dd");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("fonts");
            writer.WriteAttributeString("count", "1");
            writer.WriteStartElement("font");
            WriteValElement(writer, "sz", "11.0");
            writer.WriteStartElement("color");
            writer.WriteAttributeString("indexed", "8");
            writer.WriteEndElement();
            WriteValElement(writer, "name", "Calibri");
            WriteValElement(writer, "family", "2");
            WriteValElement(writer, "scheme", "minor");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("fills");
            writer.WriteAttributeString("count", "2");
            writer.WriteStartElement("fill");
            writer.WriteStartElement("patternFill");
            writer.WriteAttributeString("patternType", "none");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteStartElement("fill");
            writer.WriteStartElement("patternFill");
            writer.WriteAttributeString("patternType", "darkGray");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("borders");
            writer.WriteAttributeString("count", "1");
            writer.WriteStartElement("border");
            writer.WriteStartElement("left");
            writer.WriteEndElement();
            writer.WriteStartElement("right");
            writer.WriteEndElement();
            writer.WriteStartElement("top");
            writer.WriteEndElement();
            writer.WriteStartElement("bottom");
            writer.WriteEndElement();
            writer.WriteStartElement("diagonal");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("cellStyleXfs");
            writer.WriteAttributeString("count", "1");
            writer.WriteStartElement("xf");
            writer.WriteAttributeString("numFmtId", "0");
            writer.WriteAttributeString("fontId", "0");
            writer.WriteAttributeString("fillId", "0");
            writer.WriteAttributeString("borderId", "0");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("cellXfs");
            writer.WriteAttributeString("count", "2");
            writer.WriteStartElement("xf");
            writer.WriteAttributeString("numFmtId", "0");
            writer.WriteAttributeString("fontId", "0");
            writer.WriteAttributeString("fillId", "0");
            writer.WriteAttributeString("borderId", "0");
            writer.WriteAttributeString("xfId", "0");
            writer.WriteEndElement();
            writer.WriteStartElement("xf");
            writer.WriteAttributeString("numFmtId", "164");
            writer.WriteAttributeString("fontId", "0");
            writer.WriteAttributeString("fillId", "0");
            writer.WriteAttributeString("borderId", "0");
            writer.WriteAttributeString("xfId", "0");
            writer.WriteAttributeString("applyNumberFormat", "true");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("cell-types__xl__styles.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_Workbook_CellTypes_MatchesFixture()
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
            writer.WriteStartElement("sheet");
            writer.WriteAttributeString("name", "Types");
            writer.WriteAttributeString("r", "id", "rId3");
            writer.WriteAttributeString("sheetId", "1");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("cell-types__xl__workbook.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    [Fact]
    public void Write_Worksheet_CellTypes_MatchesFixture()
    {
        var builder = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(builder)))
        {
            writer.WriteStartDocument("UTF-8", standalone: false);
            writer.WriteString("\n");
            writer.WriteStartElement("worksheet");
            writer.WriteAttributeString("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            writer.WriteStartElement("dimension");
            writer.WriteAttributeString("ref", "A1:E1");
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

            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", "B1");
            writer.WriteAttributeString("t", "n");
            writer.WriteAttributeString("s", "0");
            writer.WriteStartElement("v");
            writer.WriteString("42.5");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", "C1");
            writer.WriteAttributeString("t", "b");
            writer.WriteAttributeString("s", "0");
            writer.WriteStartElement("v");
            writer.WriteString("1");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", "D1");
            writer.WriteAttributeString("t", "n");
            writer.WriteAttributeString("s", "1");
            writer.WriteStartElement("v");
            writer.WriteString("25569.375");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("c");
            writer.WriteAttributeString("r", "E1");
            writer.WriteAttributeString("s", "0");
            writer.WriteStartElement("f");
            writer.WriteString("SUM(1,2,3)");
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("pageMargins");
            writer.WriteAttributeString("bottom", "0.75");
            writer.WriteAttributeString("footer", "0.3");
            writer.WriteAttributeString("header", "0.3");
            writer.WriteAttributeString("left", "0.7");
            writer.WriteAttributeString("right", "0.7");
            writer.WriteAttributeString("top", "0.75");
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(XmlFixturePaths.GetFixturePath("cell-types__xl__worksheets__sheet1.xml"));
        Assert.Equal(expected, builder.ToString());
    }

    private static void WriteRelationship(PoiXmlWriter writer, string id, string target, string type)
    {
        writer.WriteStartElement("Relationship");
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Target", target);
        writer.WriteAttributeString("Type", type);
        writer.WriteEndElement();
    }

    private static void WriteValElement(PoiXmlWriter writer, string name, string val)
    {
        writer.WriteStartElement(name);
        writer.WriteAttributeString("val", val);
        writer.WriteEndElement();
    }
}


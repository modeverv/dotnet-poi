using System.Text;
using DotnetPoi.SS.Xml;
using Xunit;

namespace DotnetPoi.SS.Tests.Xml;

public class PoiXmlWriterContentTypesTests
{
    [Fact]
    public void Write_ContentTypes_CellEmpty_MatchesFixture()
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

            writer.WriteEndElement();
        }

        var expected = File.ReadAllText(GetFixturePath("cell-empty__[Content_Types].xml"));
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

    private static string GetFixturePath(string fileName)
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "DotnetPOI.sln")))
        {
            dir = dir.Parent;
        }

        if (dir == null)
        {
            throw new DirectoryNotFoundException("Could not locate repository root (DotnetPOI.sln).");
        }

        return Path.Combine(dir.FullName,
            "tests",
            "DotnetPoi.Interop.Tests",
            "fixtures",
            "xml-parity",
            fileName);
    }
}


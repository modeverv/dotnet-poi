using System.Text;
using DotnetPoi.SS.Xml;
using Xunit;

namespace DotnetPoi.SS.Tests.Xml;

/// <summary>
/// Verifies namespace declaration behavior matching XMLBeans/POI output (items 7 &amp; 8).
///
/// Key observations from poi-integration fixtures:
///   spreadsheet parts  → xmlns="...spreadsheetml..." xmlns:r="...relationships..."
///   drawing parts      → xmlns:xdr="...spreadsheetDrawing..." xmlns:a="...main..." xmlns:r="...relationships..."
///   No synthetic "main:" prefix appears in POI-profile output.
///
/// Implementation (item 8): PoiXmlWriter is a plain text-level writer.
/// Namespace declarations are ordinary attributes written by the caller.
/// The writer does NOT sort, hoist, or deduplicate them — callers declare
/// namespaces in POI/XMLBeans order and the writer preserves that order exactly.
/// </summary>
public class PoiXmlWriterNamespaceTests
{
    private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string SpreadsheetDrawingNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private const string DrawingMainNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

    // ── Default namespace declaration ─────────────────────────────────────

    [Fact]
    public void WriteAttributeString_DefaultNamespaceDeclaration_WritesXmlnsAttribute()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("workbook");
            writer.WriteAttributeString("xmlns", SpreadsheetNs);
            writer.WriteEndElement();
        }

        Assert.Equal($"<workbook xmlns=\"{SpreadsheetNs}\"/>", sb.ToString());
    }

    // ── r prefix declaration ──────────────────────────────────────────────

    [Fact]
    public void WriteAttributeString_RelationshipPrefixDeclaration_WritesXmlnsRAttribute()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("workbook");
            writer.WriteAttributeString("xmlns:r", RelationshipsNs);
            writer.WriteEndElement();
        }

        Assert.Equal($"<workbook xmlns:r=\"{RelationshipsNs}\"/>", sb.ToString());
    }

    [Fact]
    public void WriteAttributeString_RelationshipPrefixDeclaration_PrefixOverload_WritesXmlnsRAttribute()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("workbook");
            writer.WriteAttributeString("xmlns", "r", RelationshipsNs);
            writer.WriteEndElement();
        }

        Assert.Equal($"<workbook xmlns:r=\"{RelationshipsNs}\"/>", sb.ToString());
    }

    // ── Spreadsheet workbook root pattern: default ns + r ─────────────────

    /// Matches the root element of xl/workbook.xml observed in poi-integration fixtures:
    ///   <workbook xmlns="...spreadsheetml..." xmlns:r="...relationships...">
    [Fact]
    public void WriteAttributeString_SpreadsheetNamespacePlusR_MatchesPoiWorkbookPattern()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("workbook");
            writer.WriteAttributeString("xmlns", SpreadsheetNs);
            writer.WriteAttributeString("xmlns:r", RelationshipsNs);
            writer.WriteStartElement("sheets");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        var expected = $"<workbook xmlns=\"{SpreadsheetNs}\" xmlns:r=\"{RelationshipsNs}\"><sheets/></workbook>";
        Assert.Equal(expected, sb.ToString());
    }

    // ── Drawing root pattern: xdr + a + r ────────────────────────────────

    /// Matches the root element of xl/drawings/drawing1.xml observed in poi-integration fixtures:
    ///   <xdr:wsDr xmlns:xdr="...spreadsheetDrawing..." xmlns:a="...main..." xmlns:r="...relationships...">
    [Fact]
    public void WriteAttributeString_DrawingNamespaces_MatchesPoiDrawingPattern()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("xdr", "wsDr");
            writer.WriteAttributeString("xmlns:xdr", SpreadsheetDrawingNs);
            writer.WriteAttributeString("xmlns:a", DrawingMainNs);
            writer.WriteAttributeString("xmlns:r", RelationshipsNs);
            writer.WriteEndElement();
        }

        var expected =
            $"<xdr:wsDr xmlns:xdr=\"{SpreadsheetDrawingNs}\" xmlns:a=\"{DrawingMainNs}\" xmlns:r=\"{RelationshipsNs}\"/>";
        Assert.Equal(expected, sb.ToString());
    }

    // ── No synthetic main: prefix ─────────────────────────────────────────

    /// Elements written without a prefix are serialized without any synthetic prefix.
    /// The raw XMLBeans default save produces "main:" prefixes, but POI's DEFAULT_XML_OPTIONS
    /// removes them. PoiXmlWriter must never add a "main:" prefix on its own.
    [Fact]
    public void WriteStartElement_DefaultNamespaceElements_NoSyntheticMainPrefixAdded()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("worksheet");
            writer.WriteAttributeString("xmlns", SpreadsheetNs);
            writer.WriteStartElement("sheetData");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        var output = sb.ToString();
        Assert.DoesNotContain("main:", output);
        Assert.Contains("<sheetData/>", output);
    }

    /// Prefixed elements use the caller-supplied prefix, not a synthetic one.
    [Fact]
    public void WriteStartElement_PrefixedElements_UseCallerPrefixNotSynthetic()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("xdr", "wsDr");
            writer.WriteStartElement("xdr", "twoCellAnchor");
            writer.WriteAttributeString("editAs", "twoCell");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        var output = sb.ToString();
        Assert.DoesNotContain("main:", output);
        Assert.Contains("<xdr:wsDr>", output);
        Assert.Contains("<xdr:twoCellAnchor editAs=\"twoCell\"/>", output);
    }

    // ── No duplicate namespace declarations ───────────────────────────────

    /// When callers declare a namespace once on the root, it appears once.
    /// PoiXmlWriter does not re-declare or hoist namespaces automatically.
    [Fact]
    public void WriteAttributeString_CallerWritesNamespaceOnce_NoDuplicationInOutput()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("workbook");
            writer.WriteAttributeString("xmlns", SpreadsheetNs);
            writer.WriteStartElement("sheets");
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        var output = sb.ToString();
        // xmlns="..." should appear exactly once at the root
        var count = CountOccurrences(output, $"xmlns=\"{SpreadsheetNs}\"");
        Assert.Equal(1, count);
    }

    private static int CountOccurrences(string text, string pattern)
    {
        var count = 0;
        var index = 0;
        while ((index = text.IndexOf(pattern, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += pattern.Length;
        }
        return count;
    }
}

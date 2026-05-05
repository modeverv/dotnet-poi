using System.Text;
using DotnetPoi.SS.Xml;
using Xunit;

namespace DotnetPoi.SS.Tests.Xml;

/// <summary>
/// Verifies that PoiXmlWriter preserves attribute order exactly as callers write it (item 9).
///
/// POI/XMLBeans outputs attributes in schema/caller order, which does not necessarily
/// follow alphabetical order. PoiXmlWriter must not sort, reorder, or deduplicate
/// attributes — what callers write is what appears in the output.
///
/// Evidence from poi-integration fixtures:
///   poi-integration-comments-write-read__xl__worksheets__sheet1.xml:
///     pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"
///     (left/right/top/bottom/header/footer — not alphabetical)
///   poi-integration-sheet-layout__xl__worksheets__sheet1.xml:
///     pageMargins bottom="0.75" footer="0.3" header="0.3" left="0.7" right="0.7" top="0.75"
///     (alphabetical in this fixture — depends on what the caller writes)
/// Both orderings are valid; the key is that PoiXmlWriter must not impose alphabetical sorting.
/// </summary>
public class PoiXmlWriterAttributeOrderTests
{
    // ── Non-alphabetical caller order preserved ───────────────────────────

    /// POI page margins: left/right/top/bottom/header/footer — not alphabetical.
    /// Verifies that writing in non-alphabetical order yields non-alphabetical output.
    [Fact]
    public void WriteAttributeString_PageMarginsPoiOrder_OutputMatchesCallerOrder()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("pageMargins");
            writer.WriteAttributeString("left", "0.7");
            writer.WriteAttributeString("right", "0.7");
            writer.WriteAttributeString("top", "0.75");
            writer.WriteAttributeString("bottom", "0.75");
            writer.WriteAttributeString("header", "0.3");
            writer.WriteAttributeString("footer", "0.3");
            writer.WriteEndElement();
        }

        Assert.Equal(
            "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>",
            sb.ToString());
    }

    /// Writing z, a, m in that order — the output must preserve z before a before m.
    [Fact]
    public void WriteAttributeString_ReverseAlphabeticalOrder_OutputPreservesCallerOrder()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("e");
            writer.WriteAttributeString("z", "1");
            writer.WriteAttributeString("a", "2");
            writer.WriteAttributeString("m", "3");
            writer.WriteEndElement();
        }

        Assert.Equal("<e z=\"1\" a=\"2\" m=\"3\"/>", sb.ToString());
    }

    // ── POI .rels relationship attribute order ────────────────────────────

    /// POI Relationship elements use Id, Type, Target — exactly this order.
    /// Evidence: all .rels files in poi-integration fixtures.
    [Fact]
    public void WriteAttributeString_RelationshipIdTypeTarget_OutputPreservesPoiOrder()
    {
        const string worksheetType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("Relationship");
            writer.WriteAttributeString("Id", "rId1");
            writer.WriteAttributeString("Type", worksheetType);
            writer.WriteAttributeString("Target", "worksheets/sheet1.xml");
            writer.WriteEndElement();
        }

        Assert.Equal(
            $"<Relationship Id=\"rId1\" Type=\"{worksheetType}\" Target=\"worksheets/sheet1.xml\"/>",
            sb.ToString());
    }

    // ── Attribute order preserved across multiple elements ────────────────

    /// Two sibling elements each with their own attribute order — both preserved independently.
    [Fact]
    public void WriteAttributeString_MultipleElements_EachPreservesItsOwnAttributeOrder()
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("root");

            writer.WriteStartElement("first");
            writer.WriteAttributeString("z", "1");
            writer.WriteAttributeString("a", "2");
            writer.WriteEndElement();

            writer.WriteStartElement("second");
            writer.WriteAttributeString("a", "3");
            writer.WriteAttributeString("z", "4");
            writer.WriteEndElement();

            writer.WriteEndElement();
        }

        Assert.Equal("<root><first z=\"1\" a=\"2\"/><second a=\"3\" z=\"4\"/></root>", sb.ToString());
    }

    // ── Namespace declarations also follow caller order ───────────────────

    /// Namespace declarations are attributes — their order is also caller-controlled.
    /// POI drawing root: xmlns:xdr, xmlns:a, xmlns:r — in that order.
    [Fact]
    public void WriteAttributeString_NamespaceDeclarations_CallerOrderPreserved()
    {
        const string xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        const string a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("xdr", "wsDr");
            writer.WriteAttributeString("xmlns:xdr", xdr);
            writer.WriteAttributeString("xmlns:a", a);
            writer.WriteAttributeString("xmlns:r", r);
            writer.WriteEndElement();
        }

        var output = sb.ToString();
        var xdrPos = output.IndexOf("xmlns:xdr", StringComparison.Ordinal);
        var aPos = output.IndexOf("xmlns:a", StringComparison.Ordinal);
        var rPos = output.IndexOf("xmlns:r", StringComparison.Ordinal);

        Assert.True(xdrPos < aPos, "xmlns:xdr must appear before xmlns:a");
        Assert.True(aPos < rPos, "xmlns:a must appear before xmlns:r");
    }
}

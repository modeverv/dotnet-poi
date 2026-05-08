using System.Text;
using DotnetPoi.SS.Xml;
using Xunit;

namespace DotnetPoi.SS.Tests.Xml;

/// <summary>
/// Verifies that PoiXmlWriter escaping matches XMLBeans/POI output.
///
/// Evidence sources:
///   Text: tests/DotnetPoi.Interop.Tests/fixtures/xmlbeans-output/xmlbeans-shared-strings-escaping__poi-options.xml
///         Input "A&amp;B &lt;C> "quoted" 'single'" is stored verbatim; > is literal, quotes/apostrophes are literal.
///   Attributes (& and "): tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/
///         poi-integration-hyperlinks__xl__worksheets___rels__sheet1.xml.rels  (&amp; in URL)
///         poi-integration-styles-formatting__xl__styles.xml  (&quot; and \ in formatCode)
///
/// Divergences from System.Xml.XmlWriter (SXW) that require PoiXmlWriter to override:
///   - > in text content: POI = literal, SXW = &gt;  => PoiXmlWriter must produce literal >
///
/// Bugs in the initial PoiXmlWriter (fixed here; diverged from both POI and SXW):
///   - ' in attribute values: both POI and SXW produce literal '; initial PoiXmlWriter incorrectly produced &apos;
/// </summary>
public class PoiXmlWriterEscapingTests
{
    private static string WriteText(string content)
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("t");
            writer.WriteString(content);
            writer.WriteEndElement();
        }
        return sb.ToString();
    }

    private static string WriteAttr(string value)
    {
        var sb = new StringBuilder();
        using (var writer = new PoiXmlWriter(new StringWriter(sb)))
        {
            writer.WriteStartElement("t");
            writer.WriteAttributeString("v", value);
            writer.WriteEndElement();
        }
        return sb.ToString();
    }

    // ── Text content escaping ─────────────────────────────────────────────

    [Fact]
    public void WriteString_AmpersandInText_EscapesAsAmpEntity()
    {
        Assert.Equal("<t>A&amp;B</t>", WriteText("A&B"));
    }

    [Fact]
    public void WriteString_LessThanInText_EscapesAsLtEntity()
    {
        Assert.Equal("<t>a &lt; b</t>", WriteText("a < b"));
    }

    /// XMLBeans observation: > remains literal in text content.
    /// SXW produces &gt; — this is the proven divergence PoiXmlWriter must override.
    [Fact]
    public void WriteString_GreaterThanInText_WritesLiteralGreaterThan()
    {
        Assert.Equal("<t>a > b</t>", WriteText("a > b"));
    }

    /// XMLBeans observation: double quote is literal in element text.
    [Fact]
    public void WriteString_DoubleQuoteInText_WritesLiteralDoubleQuote()
    {
        Assert.Equal("<t>say \"hi\"</t>", WriteText("say \"hi\""));
    }

    /// XMLBeans observation: apostrophe is literal in element text.
    [Fact]
    public void WriteString_ApostropheInText_WritesLiteralApostrophe()
    {
        Assert.Equal("<t>it's</t>", WriteText("it's"));
    }

    [Fact]
    public void WriteString_TabInText_WritesLiteralTab()
    {
        Assert.Equal("<t>a\tb</t>", WriteText("a\tb"));
    }

    [Fact]
    public void WriteString_NewlineInText_WritesLiteralNewline()
    {
        Assert.Equal("<t>a\nb</t>", WriteText("a\nb"));
    }

    /// Combined case matching the XMLBeans escaping fixture:
    ///   xmlbeans-shared-strings-escaping__poi-options.xml text node = A&amp;B &lt;C> "quoted" 'single'
    [Fact]
    public void WriteString_MixedSpecialChars_MatchesXmlBeansObservation()
    {
        Assert.Equal("<t>A&amp;B &lt;C> \"quoted\" 'single'</t>",
            WriteText("A&B <C> \"quoted\" 'single'"));
    }

    // ── Attribute value escaping ──────────────────────────────────────────

    [Fact]
    public void WriteAttributeString_AmpersandInAttribute_EscapesAsAmpEntity()
    {
        Assert.Equal("<t v=\"a&amp;b\"/>", WriteAttr("a&b"));
    }

    /// POI fixture: formatCode="&quot;£&quot;#,##0.00" from poi-integration-styles-formatting.
    [Fact]
    public void WriteAttributeString_DoubleQuoteInAttribute_EscapesAsQuotEntity()
    {
        Assert.Equal("<t v=\"&quot;£&quot;#,##0.00\"/>", WriteAttr("\"£\"#,##0.00"));
    }

    /// POI fixture: formatCode="yyyy\\-mm\\-dd" — backslash is literal.
    [Fact]
    public void WriteAttributeString_BackslashInAttribute_WritesLiteralBackslash()
    {
        Assert.Equal("<t v=\"yyyy\\-mm\\-dd\"/>", WriteAttr("yyyy\\-mm\\-dd"));
    }

    /// Both POI and SXW leave ' literal in double-quoted attributes.
    /// Initial PoiXmlWriter incorrectly produced &apos; — this test captures the fix.
    [Fact]
    public void WriteAttributeString_ApostropheInAttribute_WritesLiteralApostrophe()
    {
        Assert.Equal("<t v=\"it's\"/>", WriteAttr("it's"));
    }

    [Fact]
    public void WriteAttributeString_LessThanInAttribute_EscapesAsLtEntity()
    {
        Assert.Equal("<t v=\"a &lt; b\"/>", WriteAttr("a < b"));
    }

    /// POI fixture: Target="http://apache.org/default.php?s=isTramsformed&amp;submit=Search..."
    [Fact]
    public void WriteAttributeString_RelationshipTargetWithAmpersand_EscapesAmpersand()
    {
        const string url = "http://apache.org/default.php?s=isTramsformed&submit=Search&la=*&li=*";
        const string expected = "http://apache.org/default.php?s=isTramsformed&amp;submit=Search&amp;la=*&amp;li=*";
        Assert.Equal($"<t v=\"{expected}\"/>", WriteAttr(url));
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DotnetPoi.SS.Xml;

public enum PoiXmlDeclarationProfile
{
    XmlBeansSpreadsheetPart,
    OpcPackagePart
}

public sealed class PoiXmlWriter : IDisposable
{
    private readonly TextWriter _writer;
    private readonly Stack<ElementState> _stack = new();
    private bool _startTagOpen;
    private bool _wroteDeclaration;

    public PoiXmlWriter(TextWriter writer)
    {
        _writer = writer ?? throw new ArgumentNullException(nameof(writer));
    }

    public void WriteStartDocument(PoiXmlDeclarationProfile profile)
    {
        switch (profile)
        {
            case PoiXmlDeclarationProfile.XmlBeansSpreadsheetPart:
                WriteStartDocument("UTF-8", standalone: false);
                _writer.Write('\n');
                break;
            case PoiXmlDeclarationProfile.OpcPackagePart:
                WriteStartDocument("UTF-8", standalone: true);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(profile), profile, null);
        }
    }

    public void WriteStartDocument(string encoding, bool standalone)
    {
        if (_wroteDeclaration)
        {
            throw new InvalidOperationException("XML declaration has already been written.");
        }

        _writer.Write("<?xml version=\"1.0\" encoding=\"");
        _writer.Write(encoding);
        _writer.Write("\"");
        if (standalone)
        {
            _writer.Write(" standalone=\"yes\"");
        }
        _writer.Write("?>");
        _wroteDeclaration = true;
    }

    public void WriteStartElement(string localName)
    {
        WriteStartElement(string.Empty, localName);
    }

    public void WriteStartElement(string? prefix, string localName)
    {
        CloseStartTagIfNeeded(markParentHasContent: true);

        _writer.Write('<');
        if (!string.IsNullOrEmpty(prefix))
        {
            _writer.Write(prefix);
            _writer.Write(':');
        }
        _writer.Write(localName);

        _stack.Push(new ElementState(prefix ?? string.Empty, localName));
        _startTagOpen = true;
    }

    public void WriteAttributeString(string name, string value)
    {
        if (!_startTagOpen)
        {
            throw new InvalidOperationException("Attributes can only be written immediately after a start element.");
        }

        _writer.Write(' ');
        _writer.Write(name);
        _writer.Write("=\"");
        _writer.Write(EscapeAttribute(value));
        _writer.Write('"');
    }

    public void WriteAttributeString(string? prefix, string localName, string value)
    {
        var name = string.IsNullOrEmpty(prefix) ? localName : string.Concat(prefix, ":", localName);
        WriteAttributeString(name, value);
    }

    public void WriteString(string text)
    {
        if (text.Length == 0)
        {
            return;
        }

        CloseStartTagIfNeeded(markParentHasContent: false);
        _writer.Write(EscapeText(text));
        if (_stack.Count > 0)
        {
            var state = _stack.Pop();
            state.HasContent = true;
            _stack.Push(state);
        }
    }

    public void WriteEndElement()
    {
        if (_stack.Count == 0)
        {
            throw new InvalidOperationException("No open element to close.");
        }

        var state = _stack.Pop();
        if (_startTagOpen)
        {
            _writer.Write("/>");
            _startTagOpen = false;
            return;
        }

        _writer.Write("</");
        if (!string.IsNullOrEmpty(state.Prefix))
        {
            _writer.Write(state.Prefix);
            _writer.Write(':');
        }
        _writer.Write(state.LocalName);
        _writer.Write('>');
    }

    /// <summary>
    /// Writes raw XML markup verbatim. The caller is responsible for well-formedness.
    /// </summary>
    public void WriteRaw(string data)
    {
        if (string.IsNullOrEmpty(data)) return;
        CloseStartTagIfNeeded(markParentHasContent: true);
        _writer.Write(data);
    }

    public void Dispose()
    {
        _writer.Flush();
    }

    private void CloseStartTagIfNeeded(bool markParentHasContent)
    {
        if (!_startTagOpen)
        {
            return;
        }

        _writer.Write('>');
        _startTagOpen = false;
        if (markParentHasContent && _stack.Count > 0)
        {
            var state = _stack.Pop();
            state.HasContent = true;
            _stack.Push(state);
        }
    }

    private static string EscapeAttribute(string value)
    {
        return EscapeCore(value, forAttribute: true);
    }

    private static string EscapeText(string value)
    {
        return EscapeCore(value, forAttribute: false);
    }

    private static string EscapeCore(string value, bool forAttribute)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        var builder = new StringBuilder(value.Length + 8);
        foreach (var ch in value)
        {
            switch (ch)
            {
                case '&':
                    builder.Append("&amp;");
                    break;
                case '<':
                    builder.Append("&lt;");
                    break;
                case '"' when forAttribute:
                    builder.Append("&quot;");
                    break;
                // '>' in text content: XMLBeans/POI observation shows > is literal in element text.
                // System.Xml.XmlWriter produces &gt; here — this is the one proven divergence we fix.
                // In attribute values (double-quoted), > is not required to be escaped by the XML spec,
                // and POI output for attributes is consistent with leaving it literal as well.
                default:
                    builder.Append(ch);
                    break;
            }
        }

        return builder.ToString();
    }

    private sealed class ElementState
    {
        public ElementState(string prefix, string localName)
        {
            Prefix = prefix;
            LocalName = localName;
        }

        public string Prefix { get; }
        public string LocalName { get; }
        public bool HasContent { get; set; }
    }
}

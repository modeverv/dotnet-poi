using System;
using System.IO;

namespace DotnetPoi.SS.Xml;

public static class PoiXmlWriterFactory
{
    public static PoiXmlWriter Create(TextWriter writer, PoiXmlDeclarationProfile profile)
    {
        var xmlWriter = new PoiXmlWriter(writer);
        xmlWriter.WriteStartDocument(profile);
        return xmlWriter;
    }

    public static PoiXmlWriter CreateForOoxmlPackagePart(TextWriter writer, string partName)
    {
        return Create(writer, GetDeclarationProfileForOoxmlPackagePart(partName));
    }

    public static PoiXmlDeclarationProfile GetDeclarationProfileForOoxmlPackagePart(string partName)
    {
        if (partName is null)
        {
            throw new ArgumentNullException(nameof(partName));
        }

        var normalizedName = partName.Replace('\\', '/').TrimStart('/');
        if (string.Equals(normalizedName, "[Content_Types].xml", StringComparison.Ordinal) ||
            normalizedName.EndsWith(".rels", StringComparison.Ordinal) ||
            string.Equals(normalizedName, "docProps/core.xml", StringComparison.Ordinal))
        {
            return PoiXmlDeclarationProfile.OpcPackagePart;
        }

        return PoiXmlDeclarationProfile.XmlBeansSpreadsheetPart;
    }
}

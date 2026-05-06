#!/usr/bin/env dotnet-script

// Quick verification of which features survive round-trip
// when loaded from real POI test resource files.

#r "nuget: System.IO.Compression, 4.3.0"
#r "../src/DotnetPoi.Core/bin/Release/net8.0/DotnetPoi.Core.dll"

using System.IO.Compression;
using DotnetPoi.XWPF.UserModel;
using DotnetPoi.XSSF.UserModel;
using DotnetPoi.XSLF.UserModel;

void VerifyDocx(string label, string path)
{
    if (!File.Exists(path)) { Console.WriteLine($"  SKIP — {label}: file not found"); return; }
    
    byte[] raw = File.ReadAllBytes(path);
    var beforeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    using (var z = new ZipArchive(new MemoryStream(raw), ZipArchiveMode.Read))
        foreach (var e in z.Entries) beforeNames.Add(e.FullName);

    using var doc = new XWPFDocument(new MemoryStream(raw));
    using var ms = new MemoryStream();
    doc.write(ms);
    ms.Position = 0;

    var afterNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    using (var z = new ZipArchive(ms, ZipArchiveMode.Read))
        foreach (var e in z.Entries) afterNames.Add(e.FullName);

    var lost = beforeNames.Except(afterNames).OrderBy(x => x).ToArray();
    var gained = afterNames.Except(beforeNames).OrderBy(x => x).ToArray();
    
    Console.WriteLine($"\n  {label} ({Path.GetFileName(path)})");
    if (lost.Length == 0) Console.WriteLine("    ✅ All original parts preserved");
    else foreach (var n in lost) Console.WriteLine($"    ❌ LOST: {n}");
    if (gained.Length > 0) foreach (var n in gained) Console.WriteLine($"    ➕ NEW: {n}");
}

void VerifyXlsx(string label, string path)
{
    if (!File.Exists(path)) { Console.WriteLine($"  SKIP — {label}: file not found"); return; }
    
    byte[] raw = File.ReadAllBytes(path);
    var beforeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    using (var z = new ZipArchive(new MemoryStream(raw), ZipArchiveMode.Read))
        foreach (var e in z.Entries) beforeNames.Add(e.FullName);

    using var doc = new XSSFWorkbook(new MemoryStream(raw));
    using var ms = new MemoryStream();
    doc.write(ms);
    ms.Position = 0;

    var afterNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    using (var z = new ZipArchive(ms, ZipArchiveMode.Read))
        foreach (var e in z.Entries) afterNames.Add(e.FullName);

    var lost = beforeNames.Except(afterNames).OrderBy(x => x).ToArray();
    var gained = afterNames.Except(beforeNames).OrderBy(x => x).ToArray();
    
    Console.WriteLine($"\n  {label} ({Path.GetFileName(path)})");
    if (lost.Length == 0) Console.WriteLine("    ✅ All original parts preserved");
    else foreach (var n in lost) Console.WriteLine($"    ❌ LOST: {n}");
    if (gained.Length > 0) foreach (var n in gained) Console.WriteLine($"    ➕ NEW: {n}");
}

void VerifyPptx(string label, string path)
{
    if (!File.Exists(path)) { Console.WriteLine($"  SKIP — {label}: file not found"); return; }
    
    byte[] raw = File.ReadAllBytes(path);
    var beforeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    using (var z = new ZipArchive(new MemoryStream(raw), ZipArchiveMode.Read))
        foreach (var e in z.Entries) beforeNames.Add(e.FullName);

    using var doc = new XMLSlideShow(new MemoryStream(raw));
    using var ms = new MemoryStream();
    doc.write(ms);
    ms.Position = 0;

    var afterNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    using (var z = new ZipArchive(ms, ZipArchiveMode.Read))
        foreach (var e in z.Entries) afterNames.Add(e.FullName);

    var lost = beforeNames.Except(afterNames).OrderBy(x => x).ToArray();
    var gained = afterNames.Except(beforeNames).OrderBy(x => x).ToArray();
    
    Console.WriteLine($"\n  {label} ({Path.GetFileName(path)})");
    if (lost.Length == 0) Console.WriteLine("    ✅ All original parts preserved");
    else foreach (var n in lost) Console.WriteLine($"    ❌ LOST: {n}");
    if (gained.Length > 0) foreach (var n in gained) Console.WriteLine($"    ➕ NEW: {n}");
}

Console.WriteLine("═══ DOCX round-trip preservation test ═══");
VerifyDocx("Styles+comments test", "/workspace/poi/test-data/document/55966.docx");
VerifyDocx("Simple docx", "/workspace/poi/test-data/document/56392.docx");
VerifyDocx("Another docx", "/workspace/poi/test-data/document/57312.docx");

Console.WriteLine("\n═══ XLSX round-trip preservation test ═══");
VerifyXlsx("Charts+drawings xlsx", "/workspace/poi/test-data/spreadsheet/123233_charts.xlsx");
VerifyXlsx("Footer xlsx", "/workspace/poi/test-data/spreadsheet/45540_classic_Footer.xlsx");
VerifyXlsx("45544 xlsx", "/workspace/poi/test-data/spreadsheet/45544.xlsx");

Console.WriteLine("\n═══ PPTX round-trip preservation test ═══");
VerifyPptx("Performance pptx", "/workspace/poi/test-data/slideshow/2411-Performance_Up.pptx");
VerifyPptx("Footer pptx", "/workspace/poi/test-data/slideshow/45541_Footer.pptx");
VerifyPptx("Comment pptx", "/workspace/poi/test-data/slideshow/45545_Comment.pptx");

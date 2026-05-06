using System.IO.Compression;
using DotnetPoi.XWPF.UserModel;

var path = "/workspace/poi/test-data/document/60316.docx";
var raw = File.ReadAllBytes(path);

using var doc = new XWPFDocument(new MemoryStream(raw));

using var ms = new MemoryStream();
doc.write(ms);
var output = ms.ToArray();
using var archive = new ZipArchive(new MemoryStream(output), ZipArchiveMode.Read);
var entry = archive.GetEntry("word/document.xml");
if (entry == null) { Console.WriteLine("ERROR: word/document.xml not found!"); return; }
using var r = new StreamReader(entry.Open());
var xml = r.ReadToEnd();
int sdtCount = 0, idx = 0;
while ((idx = xml.IndexOf("<w:sdt", idx, StringComparison.Ordinal)) != -1) { sdtCount++; idx++; }
Console.WriteLine($"Output <w:sdt> count: {sdtCount}");
Console.WriteLine($"Contains txbxContent: {xml.Contains("txbxContent", StringComparison.Ordinal)}");

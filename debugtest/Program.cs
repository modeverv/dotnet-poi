using System.Xml;
using System.IO.Compression;
using DotnetPoi.XWPF.UserModel;

// Create a docx
var doc = new XWPFDocument();
var para = doc.createParagraph();
var run = para.createRun();
run.setText("styled text");
run.setFontName("Arial");
run.setFontSize(12.0);
run.setColor("FF0000");
para.setAlignment(ParagraphAlignment.Center);

using var stream = new MemoryStream();
doc.write(stream);
stream.Position = 0;

// Read and trace
using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
var entry = archive.GetEntry("word/document.xml");
using var rdr = XmlReader.Create(entry.Open());
const string NsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
bool inRPr = false;
bool inPPr = false;

while (rdr.Read())
{
    if (rdr.NodeType == XmlNodeType.Element && rdr.NamespaceURI == NsW)
    {
        Console.WriteLine($"ELEMENT: <{rdr.LocalName}> inRPr={inRPr} inPPr={inPPr}");
        
        // Show all attributes
        if (rdr.HasAttributes)
        {
            while (rdr.MoveToNextAttribute())
            {
                Console.WriteLine($"  ATTR: {rdr.Name} = '{rdr.Value}' (local={rdr.LocalName}, ns='{rdr.NamespaceURI}')");
            }
            rdr.MoveToElement();
        }

        if (rdr.LocalName == "rPr") inRPr = true;
        if (rdr.LocalName == "pPr") inPPr = true;
        if (rdr.LocalName == "rFonts" && inRPr)
        {
            Console.WriteLine($"  rFonts.ascii via GetAttribute('ascii'): '{rdr.GetAttribute("ascii")}'");
            Console.WriteLine($"  rFonts.ascii via GetAttribute('w:ascii'): '{rdr.GetAttribute("w:ascii")}'");
        }
        if (rdr.LocalName == "sz" && inRPr)
        {
            Console.WriteLine($"  sz.val via GetAttribute('val'): '{rdr.GetAttribute("val")}'");
        }
        if (rdr.LocalName == "color" && inRPr)
        {
            Console.WriteLine($"  color.val via GetAttribute('val'): '{rdr.GetAttribute("val")}'");
        }
        if (rdr.LocalName == "jc" && inPPr)
        {
            Console.WriteLine($"  jc.val via GetAttribute('val'): '{rdr.GetAttribute("val")}'");
        }
    }
    if (rdr.NodeType == XmlNodeType.EndElement && rdr.NamespaceURI == NsW)
    {
        if (rdr.LocalName == "rPr") inRPr = false;
        if (rdr.LocalName == "pPr") inPPr = false;
    }
}

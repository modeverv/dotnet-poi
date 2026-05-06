using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

var sourceRoot = args.Length > 0 ? Path.GetFullPath(args[0]) : Path.GetFullPath("docs_src");
var outputRoot = args.Length > 1 ? Path.GetFullPath(args[1]) : Path.GetFullPath("docs");

if (!Directory.Exists(sourceRoot))
{
    Console.Error.WriteLine($"Docs source directory not found: {sourceRoot}");
    return 1;
}

var contentRoot = Path.Combine(sourceRoot, "content");
if (!Directory.Exists(contentRoot))
{
    Console.Error.WriteLine($"Docs content directory not found: {contentRoot}");
    return 1;
}

Directory.CreateDirectory(outputRoot);

var config = LoadSiteConfig(Path.Combine(sourceRoot, "site.json"));
var pages = LoadPages(contentRoot, outputRoot);

if (pages.Count == 0)
{
    Console.Error.WriteLine($"No Markdown pages found under: {contentRoot}");
    return 1;
}

CopyAssets(Path.Combine(sourceRoot, "assets"), Path.Combine(outputRoot, "assets"));
WritePages(config, pages, outputRoot);
WriteIndex(config, pages, outputRoot);

Console.WriteLine($"Generated {pages.Count + 1} HTML files in {outputRoot}");
return 0;

static SiteConfig LoadSiteConfig(string path)
{
    if (!File.Exists(path))
        return new SiteConfig("dotnet-poi Documentation", "Usage-focused documentation.", []);

    var json = File.ReadAllText(path);
    return JsonSerializer.Deserialize<SiteConfig>(json, new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    }) ?? new SiteConfig("dotnet-poi Documentation", "Usage-focused documentation.", []);
}

static List<DocPage> LoadPages(string contentRoot, string outputRoot)
{
    var pages = new List<DocPage>();
    foreach (var markdownPath in Directory.EnumerateFiles(contentRoot, "*.md", SearchOption.AllDirectories).OrderBy(p => p, StringComparer.Ordinal))
    {
        var relativeMarkdown = Path.GetRelativePath(contentRoot, markdownPath).Replace('\\', '/');
        var relativeHtml = Path.ChangeExtension(relativeMarkdown, ".html");
        var outputPath = Path.Combine(outputRoot, relativeHtml.Replace('/', Path.DirectorySeparatorChar));
        var markdown = File.ReadAllText(markdownPath);
        var title = ExtractTitle(markdown) ?? Path.GetFileNameWithoutExtension(markdownPath);
        var html = Markdown.ToHtml(markdown);
        pages.Add(new DocPage(title, relativeMarkdown, relativeHtml, outputPath, html));
    }

    return pages;
}

static string? ExtractTitle(string markdown)
{
    using var reader = new StringReader(markdown);
    string? line;
    while ((line = reader.ReadLine()) is not null)
    {
        if (line.StartsWith("# ", StringComparison.Ordinal))
            return line[2..].Trim();
    }

    return null;
}

static void CopyAssets(string sourceAssets, string outputAssets)
{
    if (!Directory.Exists(sourceAssets))
        return;

    foreach (var file in Directory.EnumerateFiles(sourceAssets, "*", SearchOption.AllDirectories))
    {
        var relative = Path.GetRelativePath(sourceAssets, file);
        var destination = Path.Combine(outputAssets, relative);
        Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        File.Copy(file, destination, overwrite: true);
    }
}

static void WritePages(SiteConfig config, IReadOnlyList<DocPage> pages, string outputRoot)
{
    foreach (var page in pages)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(page.OutputPath)!);
        File.WriteAllText(page.OutputPath, Layout(config, pages, page, page.BodyHtml, outputRoot), new UTF8Encoding(false));
    }
}

static void WriteIndex(SiteConfig config, IReadOnlyList<DocPage> pages, string outputRoot)
{
    var body = new StringBuilder();
    body.AppendLine("<h1>dotnet-poi Documentation</h1>");
    body.AppendLine($"<p>{Html(config.Description)}</p>");
    body.AppendLine("<h2>Start Here</h2>");
    body.AppendLine("<ul>");
    foreach (var page in pages)
    {
        body.Append("<li><a href=\"");
        body.Append(Html(page.RelativeHtml));
        body.Append("\">");
        body.Append(Html(page.Title));
        body.AppendLine("</a></li>");
    }
    body.AppendLine("</ul>");

    var indexPage = new DocPage("Home", "index.md", "index.html", Path.Combine(outputRoot, "index.html"), body.ToString());
    File.WriteAllText(indexPage.OutputPath, Layout(config, pages, indexPage, body.ToString(), outputRoot), new UTF8Encoding(false));
}

static string Layout(SiteConfig config, IReadOnlyList<DocPage> pages, DocPage current, string body, string outputRoot)
{
    var currentDirectory = Path.GetDirectoryName(current.OutputPath) ?? outputRoot;
    var assetHref = RelativeHref(currentDirectory, Path.Combine(outputRoot, "assets", "site.css"));

    var nav = new StringBuilder();
    foreach (var section in ResolveNavigation(config, pages))
    {
        nav.AppendLine("<div class=\"nav-section\">");
        nav.Append("<div class=\"nav-title\">");
        nav.Append(Html(section.Title));
        nav.AppendLine("</div>");
        nav.AppendLine("<ul>");
        foreach (var item in section.Items)
        {
            var page = pages.FirstOrDefault(p => string.Equals(p.RelativeMarkdown, item.Path, StringComparison.OrdinalIgnoreCase))
                ?? pages.FirstOrDefault(p => string.Equals(p.RelativeHtml, Path.ChangeExtension(item.Path, ".html"), StringComparison.OrdinalIgnoreCase));
            if (page is null)
                continue;

            var href = RelativeHref(currentDirectory, page.OutputPath);
            var active = string.Equals(page.RelativeHtml, current.RelativeHtml, StringComparison.OrdinalIgnoreCase)
                ? " class=\"active\""
                : string.Empty;
            nav.Append("<li><a");
            nav.Append(active);
            nav.Append(" href=\"");
            nav.Append(Html(href));
            nav.Append("\">");
            nav.Append(Html(item.Title));
            nav.AppendLine("</a></li>");
        }
        nav.AppendLine("</ul>");
        nav.AppendLine("</div>");
    }

    if (nav.Length == 0)
    {
        nav.AppendLine("<div class=\"nav-section\"><div class=\"nav-title\">Pages</div><ul>");
        foreach (var page in pages)
        {
            nav.Append("<li><a href=\"");
            nav.Append(Html(RelativeHref(currentDirectory, page.OutputPath)));
            nav.Append("\">");
            nav.Append(Html(page.Title));
            nav.AppendLine("</a></li>");
        }
        nav.AppendLine("</ul></div>");
    }

    return $$"""
        <!doctype html>
        <html lang="en">
        <head>
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>{{Html(current.Title)}} | {{Html(config.Title)}}</title>
          <meta name="description" content="{{Html(config.Description)}}">
          <link rel="stylesheet" href="{{Html(assetHref)}}">
        </head>
        <body>
          <div class="site-shell">
            <aside class="sidebar">
              <a class="brand" href="{{Html(RelativeHref(currentDirectory, Path.Combine(outputRoot, "index.html")))}}">{{Html(config.Title)}}</a>
              <nav>
        {{nav.ToString().TrimEnd()}}
              </nav>
            </aside>
            <main class="content">
        {{body}}
            </main>
          </div>
        </body>
        </html>
        """;
}

static IReadOnlyList<NavSection> ResolveNavigation(SiteConfig config, IReadOnlyList<DocPage> pages)
{
    if (config.Navigation.Count > 0)
        return config.Navigation;

    return
    [
        new NavSection("Pages", pages.Select(p => new NavItem(p.Title, p.RelativeMarkdown)).ToList())
    ];
}

static string RelativeHref(string fromDirectory, string toPath)
{
    var relative = Path.GetRelativePath(fromDirectory, toPath).Replace('\\', '/');
    return relative == "." ? "index.html" : relative;
}

static string Html(string value) => WebUtility.HtmlEncode(value);

internal sealed record SiteConfig(string Title, string Description, List<NavSection> Navigation);
internal sealed record NavSection(string Title, List<NavItem> Items);
internal sealed record NavItem(string Title, string Path);
internal sealed record DocPage(string Title, string RelativeMarkdown, string RelativeHtml, string OutputPath, string BodyHtml);

internal static partial class Markdown
{
    public static string ToHtml(string markdown)
    {
        var html = new StringBuilder();
        var paragraph = new StringBuilder();
        var listOpen = false;
        var tableOpen = false;
        var inCode = false;
        var codeLanguage = string.Empty;
        var code = new StringBuilder();

        foreach (var rawLine in markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
        {
            var line = rawLine.TrimEnd();

            if (line.StartsWith("```", StringComparison.Ordinal))
            {
                if (inCode)
                {
                    html.Append("<pre><code");
                    if (!string.IsNullOrWhiteSpace(codeLanguage))
                    {
                        html.Append(" class=\"language-");
                        html.Append(WebUtility.HtmlEncode(codeLanguage));
                        html.Append('"');
                    }
                    html.Append('>');
                    html.Append(WebUtility.HtmlEncode(code.ToString().TrimEnd('\n')));
                    html.AppendLine("</code></pre>");
                    code.Clear();
                    codeLanguage = string.Empty;
                    inCode = false;
                }
                else
                {
                    FlushParagraph(html, paragraph);
                    CloseList(html, ref listOpen);
                    CloseTable(html, ref tableOpen);
                    codeLanguage = line[3..].Trim();
                    inCode = true;
                }
                continue;
            }

            if (inCode)
            {
                code.AppendLine(rawLine);
                continue;
            }

            if (string.IsNullOrWhiteSpace(line))
            {
                FlushParagraph(html, paragraph);
                CloseList(html, ref listOpen);
                CloseTable(html, ref tableOpen);
                continue;
            }

            // Table row: must start and end with |
            if (line.StartsWith('|') && line.EndsWith('|'))
            {
                FlushParagraph(html, paragraph);
                CloseList(html, ref listOpen);

                // Separator row (|---|) — skip but ensure table is open
                if (line.Length > 2 && line.All(c => c is '|' or '-' or ':' or ' '))
                {
                    if (!tableOpen)
                    {
                        html.AppendLine("<table>");
                        tableOpen = true;
                    }
                    continue;
                }

                if (!tableOpen)
                {
                    html.AppendLine("<table>");
                    tableOpen = true;

                    // First data row is treated as header
                    html.Append("<thead><tr>");
                    var headerCells = line.Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                    foreach (var cell in headerCells)
                    {
                        html.Append("<th>");
                        html.Append(InlineWithBold(cell));
                        html.Append("</th>");
                    }
                    html.AppendLine("</tr></thead>");
                    html.AppendLine("<tbody>");
                }
                else
                {
                    html.Append("<tr>");
                    var cells = line.Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                    foreach (var cell in cells)
                    {
                        html.Append("<td>");
                        html.Append(InlineWithBold(cell));
                        html.Append("</td>");
                    }
                    html.AppendLine("</tr>");
                }
                continue;
            }

            if (line.StartsWith("#", StringComparison.Ordinal))
            {
                var level = line.TakeWhile(c => c == '#').Count();
                if (level is >= 1 and <= 6 && line.Length > level && line[level] == ' ')
                {
                    FlushParagraph(html, paragraph);
                    CloseList(html, ref listOpen);
                    CloseTable(html, ref tableOpen);
                    var text = line[(level + 1)..].Trim();
                    html.Append("<h");
                    html.Append(level);
                    html.Append('>');
                    html.Append(InlineWithBold(text));
                    html.Append("</h");
                    html.Append(level);
                    html.AppendLine(">");
                    continue;
                }
            }

            if (line.StartsWith("- ", StringComparison.Ordinal))
            {
                FlushParagraph(html, paragraph);
                CloseTable(html, ref tableOpen);
                if (!listOpen)
                {
                    html.AppendLine("<ul>");
                    listOpen = true;
                }
                html.Append("<li>");
                html.Append(InlineWithBold(line[2..].Trim()));
                html.AppendLine("</li>");
                continue;
            }

            if (paragraph.Length > 0)
                paragraph.Append(' ');
            paragraph.Append(line.Trim());
        }

        if (inCode)
        {
            html.Append("<pre><code>");
            html.Append(WebUtility.HtmlEncode(code.ToString().TrimEnd('\n')));
            html.AppendLine("</code></pre>");
        }

        FlushParagraph(html, paragraph);
        CloseList(html, ref listOpen);
        CloseTable(html, ref tableOpen);
        return html.ToString();
    }

    private static void FlushParagraph(StringBuilder html, StringBuilder paragraph)
    {
        if (paragraph.Length == 0)
            return;

        html.Append("<p>");
        html.Append(InlineWithBold(paragraph.ToString()));
        html.AppendLine("</p>");
        paragraph.Clear();
    }

    private static void CloseList(StringBuilder html, ref bool listOpen)
    {
        if (!listOpen)
            return;

        html.AppendLine("</ul>");
        listOpen = false;
    }

    private static void CloseTable(StringBuilder html, ref bool tableOpen)
    {
        if (!tableOpen)
            return;

        html.AppendLine("</tbody>");
        html.AppendLine("</table>");
        tableOpen = false;
    }

    private static string InlineWithBold(string value)
    {
        var encoded = WebUtility.HtmlEncode(value);

        encoded = BoldPattern().Replace(encoded, match => $"<strong>{match.Groups[1].Value}</strong>");

        encoded = LinkPattern().Replace(encoded, match =>
        {
            var text = match.Groups[1].Value;
            var url = match.Groups[2].Value;
            return $"<a href=\"{url}\">{text}</a>";
        });

        encoded = CodePattern().Replace(encoded, match => $"<code>{match.Groups[1].Value}</code>");
        return encoded;
    }

    [GeneratedRegex(@"\*\*([^*]+)\*\*")]
    private static partial Regex BoldPattern();

    [GeneratedRegex(@"\[([^\]]+)\]\(([^)]+)\)")]
    private static partial Regex LinkPattern();

    [GeneratedRegex("`([^`]+)`")]
    private static partial Regex CodePattern();
}

using System.Net;
using System.Text;
using System.Xml.Linq;
using HtmlAgilityPack;
using OneNote2Md.ComClient;

namespace OneNote2Md.Renderer;

/// <summary>
/// Renders a OneNote page (its raw XML) into Markdown: title, headings, lists, tables, inline
/// formatting, fenced code blocks (with an inferred language), and extracted images.
/// </summary>
public sealed class MarkdownRenderer
{
    private static readonly XNamespace One = OneNoteComClient.One;

    // Per-render link rewriter; set at the start of each Render call. Rendering is sequential, so a
    // single instance field is sufficient.
    private Func<string, string?>? _linkResolver;

    // Per-render highlight style, set at the start of each Render call.
    private HighlightStyle _highlight;

    private static readonly string[] MonospaceFonts =
    {
        "consolas", "courier", "courier new", "lucida console", "lucida sans typewriter",
        "menlo", "monaco", "dejavu sans mono", "cascadia", "monospace",
    };

    /// <summary>Renders a OneNote page's raw XML into Markdown and extracts its images.</summary>
    /// <param name="pageXml">The raw OneNote page XML (fetched with binary image data).</param>
    /// <param name="ctx">Per-page rendering settings.</param>
    /// <returns>The rendered Markdown together with any images to write alongside the page.</returns>
    public RenderedPage Render(string pageXml, RenderContext ctx)
    {
        var doc = XDocument.Parse(pageXml);
        _linkResolver = ctx.LinkResolver;
        _highlight = ctx.Highlight;

        // OneNote references paragraph styles (h1..h6, normal, etc.) by index; build index -> name lookup.
        var styleMap = doc.Descendants(One + "QuickStyleDef")
            .GroupBy(d => (string?)d.Attribute("index") ?? string.Empty)
            .ToDictionary(g => g.Key, g => (string?)g.First().Attribute("name") ?? string.Empty);

        var title = doc.Descendants(One + "Title")
            .Descendants(One + "OE")
            .Descendants(One + "T")
            .FirstOrDefault()?.Value;
        var titleText = string.IsNullOrWhiteSpace(title) ? string.Empty : InlineHtmlToMarkdown(title).Trim();

        var images = new List<ExtractedImage>();
        var attachments = new List<ExtractedAttachment>();
        var state = new RenderState(ctx, styleMap, images, attachments);
        var sb = new StringBuilder();

        if (ctx.FrontMatter)
        {
            AppendFrontMatter(sb, doc, titleText, ctx);
        }

        if (ctx.TitleHeading && !string.IsNullOrEmpty(titleText))
        {
            sb.Append('#', Math.Max(1, ctx.HeadingOffset)).Append(' ').Append(ToInlineBreaks(titleText)).Append("\n\n");
        }

        foreach (var outline in doc.Descendants(One + "Outline"))
        {
            foreach (var children in outline.Elements(One + "OEChildren"))
            {
                RenderOeChildren(children, state, sb, 0);
            }
        }

        return new RenderedPage(sb.ToString().TrimEnd() + "\n", images, attachments);
    }

    private sealed class RenderState
    {
        public RenderState(RenderContext ctx, IReadOnlyDictionary<string, string> styleMap,
            List<ExtractedImage> images, List<ExtractedAttachment> attachments)
        {
            Ctx = ctx;
            StyleMap = styleMap;
            Images = images;
            Attachments = attachments;
        }

        public RenderContext Ctx { get; }
        public IReadOnlyDictionary<string, string> StyleMap { get; }
        public List<ExtractedImage> Images { get; }
        public List<ExtractedAttachment> Attachments { get; }
    }

    private void RenderOeChildren(XElement oeChildren, RenderState state, StringBuilder sb, int depth)
    {
        var items = oeChildren.Elements(One + "OE").ToList();
        for (int i = 0; i < items.Count; i++)
        {
            var oe = items[i];

            // Group consecutive monospace paragraphs into a single fenced code block.
            if (IsCodeParagraph(oe))
            {
                var lines = new List<string>();
                while (i < items.Count && IsCodeParagraph(items[i]))
                {
                    lines.Add(CodeText(items[i]));
                    i++;
                }
                i--; // step back; the for-loop will advance past the last code line

                sb.Append("```").Append(state.Ctx.CodeLanguage).Append('\n');
                sb.Append(string.Join("\n", lines).TrimEnd()).Append("\n```\n\n");
                continue;
            }

            RenderOe(oe, state, sb, depth);
        }
    }

    private void RenderOe(XElement oe, RenderState state, StringBuilder sb, int depth)
    {
        var table = oe.Element(One + "Table");
        var image = oe.Element(One + "Image");
        var insertedFile = oe.Element(One + "InsertedFile");

        if (table is not null)
        {
            EnsureBlankLine(sb);
            RenderTable(table, sb);
        }
        else if (image is not null)
        {
            EnsureBlankLine(sb);
            RenderImage(image, state, sb);
        }
        else if (insertedFile is not null)
        {
            EnsureBlankLine(sb);
            RenderInsertedFile(insertedFile, state, sb);
        }
        else
        {
            var text = string.Concat(oe.Elements(One + "T").Select(t => InlineHtmlToMarkdown(t.Value))).Trim();
            var list = oe.Element(One + "List");
            var styleIndex = (string?)oe.Attribute("quickStyleIndex");
            var styleName = styleIndex is not null && state.StyleMap.TryGetValue(styleIndex, out var n) ? n : string.Empty;

            if (list is not null)
            {
                RenderListItem(list, text, sb, depth);
            }
            else if (IsHeading(styleName, out var level))
            {
                EnsureBlankLine(sb);
                var hashes = Math.Min(level + state.Ctx.HeadingOffset, 6);
                sb.Append('#', hashes).Append(' ').Append(ToInlineBreaks(text)).Append("\n\n");
            }
            else if (!string.IsNullOrEmpty(text))
            {
                EnsureBlankLine(sb);
                sb.Append(ToParagraphBreaks(EscapeLeadingNumber(text))).Append("\n\n");
            }
        }

        var nextDepth = oe.Element(One + "List") is not null ? depth + 1 : depth;
        foreach (var children in oe.Elements(One + "OEChildren"))
        {
            RenderOeChildren(children, state, sb, nextDepth);
        }
    }

    // Spaces of indentation added per list-nesting level. Three aligns a nested item with the content
    // column of a single-digit ordered parent ("1. "), which Markdown requires to treat it as nested.
    private const int ListIndentUnit = 3;

    // Renders one list item. OneNote gives us the exact authored marker via <one:List><one:Number text>.
    // Numeric markers are emitted as real Markdown ordered items using their explicit number (so an
    // interrupting block doesn't restart the count); bullets use '-'; alpha/roman markers (which Markdown
    // cannot represent natively) are emitted as literal, hard-broken continuation lines.
    private void RenderListItem(XElement list, string text, StringBuilder sb, int depth)
    {
        var indent = new string(' ', depth * ListIndentUnit);
        var body = ToInlineBreaks(text);
        var number = list.Element(One + "Number");

        if (number is null)
        {
            sb.Append(indent).Append("- ").Append(body).Append('\n');
            return;
        }

        var marker = ((string?)number.Attribute("text") ?? string.Empty).Trim();
        if (IsNumericMarker(marker))
        {
            sb.Append(indent).Append(marker).Append(' ').Append(body).Append('\n');
        }
        else
        {
            var literal = marker.Length == 0 ? "-" : marker;
            sb.Append(indent).Append(literal).Append(' ').Append(body).Append("  \n");
        }
    }

    // A numeric ordered-list marker is one or more digits followed by '.' or ')', e.g. "1." or "2)".
    private static bool IsNumericMarker(string marker) =>
        marker.Length >= 2 &&
        (marker[^1] is '.' or ')') &&
        marker[..^1].All(char.IsDigit);

    // Escapes a paragraph that begins with what looks like an ordered-list marker ("1.", "2)") so
    // Markdown renders the literal number the author typed instead of turning it into an auto-numbered list.
    private static string EscapeLeadingNumber(string text)
    {
        var match = LeadingNumber.Match(text);
        if (!match.Success) return text;

        var digits = match.Groups[1].Value;
        var delimiter = match.Groups[2].Value;
        return $"{digits}\\{delimiter}{text[match.Length..]}";
    }

    private static readonly System.Text.RegularExpressions.Regex LeadingNumber =
        new(@"^(\d+)([.)])(?=\s)", System.Text.RegularExpressions.RegexOptions.Compiled);

    // Ensures the buffer ends with a blank line (when it already holds content), so the next block is
    // separated from a preceding list item; without it Markdown folds the block into the last bullet.
    private static void EnsureBlankLine(StringBuilder sb)
    {
        if (sb.Length == 0) return;

        if (sb[^1] != '\n')
        {
            sb.Append("\n\n");
        }
        else if (sb.Length >= 2 && sb[^2] != '\n')
        {
            sb.Append('\n');
        }
    }

    private void RenderImage(XElement image, RenderState state, StringBuilder sb)
    {
        if (state.Ctx.Images == ImageMode.Skip || state.Ctx.ImageFileAllocator is null)
        {
            return;
        }

        var data = image.Element(One + "Data")?.Value;
        if (string.IsNullOrWhiteSpace(data)) return;

        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(data.Trim());
        }
        catch (FormatException)
        {
            return;
        }

        var ext = NormalizeImageExtension((string?)image.Attribute("format"));
        var relativePath = state.Ctx.ImageFileAllocator(ext);
        state.Images.Add(new ExtractedImage(relativePath, bytes));

        var altText = EscapeLinkText(Path.GetFileNameWithoutExtension(relativePath));
        sb.Append("![").Append(altText).Append("](").Append(ToAssetLink(relativePath)).Append(")\n\n");
    }

    // Copies an embedded file (a OneNote InsertedFile) out via its local cache and links to it. The
    // bytes live in 'pathCache', which OneNote populates when the page is fetched with binary data.
    private void RenderInsertedFile(XElement file, RenderState state, StringBuilder sb)
    {
        var preferred = ((string?)file.Attribute("preferredName"))?.Trim();
        if (string.IsNullOrWhiteSpace(preferred)) return;

        var cache = (string?)file.Attribute("pathCache");
        if (state.Ctx.AttachmentFileAllocator is null ||
            string.IsNullOrWhiteSpace(cache) || !File.Exists(cache))
        {
            // The file was never cached locally (e.g. an un-synced page); keep a visible marker
            // rather than emitting a link to a file we cannot write.
            sb.Append("**[missing attachment: ").Append(EscapeLinkText(preferred)).Append("]**\n\n");
            return;
        }

        var relativePath = state.Ctx.AttachmentFileAllocator(preferred);
        state.Attachments.Add(new ExtractedAttachment(relativePath, cache));
        sb.Append('[').Append(EscapeLinkText(preferred)).Append("](").Append(ToAssetLink(relativePath)).Append(")\n\n");
    }

    // Wraps a relative asset path in angle brackets when it contains spaces so the Markdown link parses.
    private static string ToAssetLink(string relativePath) =>
        relativePath.Contains(' ') ? $"<{relativePath}>" : relativePath;

    private static string EscapeLinkText(string text) => text.Replace("[", "\\[").Replace("]", "\\]");

    private void RenderTable(XElement table, StringBuilder sb)
    {
        var rows = table.Elements(One + "Row").ToList();
        if (rows.Count == 0) return;

        for (var i = 0; i < rows.Count; i++)
        {
            var cells = rows[i].Elements(One + "Cell").Select(GetCellText).ToList();
            sb.Append("| ").Append(string.Join(" | ", cells)).Append(" |\n");
            if (i == 0)
            {
                sb.Append("| ").Append(string.Join(" | ", cells.Select(_ => "---"))).Append(" |\n");
            }
        }
        sb.Append('\n');
    }

    private string GetCellText(XElement cell)
    {
        // Each OE inside a cell is its own line; preserve those breaks plus any inline <br>.
        var lines = cell.Descendants(One + "OE")
            .Select(oe => string.Concat(oe.Elements(One + "T").Select(t => InlineHtmlToMarkdown(t.Value))).Trim())
            .Where(s => s.Length > 0);
        return ToInlineBreaks(string.Join("\n", lines)).Replace("|", "\\|");
    }

    // ---- front matter ------------------------------------------------------

    private static void AppendFrontMatter(StringBuilder sb, XDocument doc, string titleText, RenderContext ctx)
    {
        var page = doc.Descendants(One + "Page").FirstOrDefault();
        var id = (string?)page?.Attribute("ID") ?? string.Empty;
        var created = (string?)page?.Attribute("dateTime");
        var modified = (string?)page?.Attribute("lastModifiedTime");

        // OneNote has no page-level author; authorship is recorded per outline element (<one:OE>).
        // Derive the page creator from the earliest-created paragraph and the last editor from the
        // most-recently-modified paragraph (falling back to its original author when OneNote did not
        // record a distinct modifier). Times are ISO-8601 UTC, so ordinal string ordering is correct.
        string? createdBy = null, modifiedBy = null;
        var authoredOes = doc.Descendants(One + "OE")
            .Where(e => e.Attribute("creationTime") != null)
            .ToList();
        if (authoredOes.Count > 0)
        {
            var firstCreated = authoredOes
                .OrderBy(e => (string?)e.Attribute("creationTime"), StringComparer.Ordinal)
                .First();
            createdBy = (string?)firstCreated.Attribute("author");

            var lastModified = authoredOes
                .OrderByDescending(e => (string?)e.Attribute("lastModifiedTime"), StringComparer.Ordinal)
                .First();
            modifiedBy = (string?)lastModified.Attribute("lastModifiedBy");
            if (string.IsNullOrEmpty(modifiedBy))
                modifiedBy = (string?)lastModified.Attribute("author");
        }

        sb.Append("---\n");
        sb.Append("title: ").Append(YamlScalar(titleText)).Append('\n');
        if (!string.IsNullOrEmpty(created)) sb.Append("created: ").Append(created).Append('\n');
        if (!string.IsNullOrEmpty(createdBy)) sb.Append("created_by: ").Append(YamlScalar(createdBy)).Append('\n');
        if (!string.IsNullOrEmpty(modified)) sb.Append("modified: ").Append(modified).Append('\n');
        if (!string.IsNullOrEmpty(modifiedBy)) sb.Append("modified_by: ").Append(YamlScalar(modifiedBy)).Append('\n');
        if (!string.IsNullOrEmpty(ctx.SourcePath)) sb.Append("source: ").Append(YamlScalar(ctx.SourcePath)).Append('\n');
        if (!string.IsNullOrEmpty(id)) sb.Append("onenote_id: ").Append(YamlScalar(id)).Append('\n');
        sb.Append("---\n\n");
    }

    private static string YamlScalar(string value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";

    // ---- code detection ----------------------------------------------------

    // A paragraph counts as code when it is plain text rendered in a monospace font and has no
    // list marker, table, image, or nested children.
    private static bool IsCodeParagraph(XElement oe)
    {
        if (oe.Element(One + "List") is not null) return false;
        if (oe.Element(One + "Table") is not null) return false;
        if (oe.Element(One + "Image") is not null) return false;
        if (oe.Elements(One + "OEChildren").Any()) return false;

        var runs = oe.Elements(One + "T").ToList();
        if (runs.Count == 0) return false;
        if (runs.All(t => string.IsNullOrWhiteSpace(HtmlToPlainText(t.Value)))) return false;

        var styleHaystack = ((string?)oe.Attribute("style") ?? string.Empty) + " " +
                            string.Concat(runs.Select(t => t.Value));
        styleHaystack = styleHaystack.ToLowerInvariant();
        return MonospaceFonts.Any(f => styleHaystack.Contains("font-family:" + f) || styleHaystack.Contains("font-family: " + f));
    }

    private static string CodeText(XElement oe) =>
        string.Join("\n", oe.Elements(One + "T").Select(t => HtmlToPlainText(t.Value)));

    // ---- helpers (ported from onenote-mcp) ---------------------------------

    private static string NormalizeImageExtension(string? format)
    {
        var f = (format ?? "png").Trim().ToLowerInvariant();
        return f switch
        {
            "jpg" or "jpeg" => "jpg",
            "gif" => "gif",
            "bmp" => "bmp",
            "tiff" or "tif" => "tif",
            "emf" => "emf",
            "wmf" => "wmf",
            _ => "png",
        };
    }

    // Markdown paragraph hard break: two trailing spaces before the newline force a real line break.
    private static string ToParagraphBreaks(string text) => NormalizeNewlines(text).Replace("\n", "  \n");

    // Line breaks that must stay on a single logical line (table cells, headings, list items) become <br>.
    private static string ToInlineBreaks(string text) => NormalizeNewlines(text).Replace("\n", "<br>");

    private static string NormalizeNewlines(string text) => text.Replace("\r\n", "\n").Replace('\r', '\n');

    // OneNote heading styles are named "h1".."h6".
    private static bool IsHeading(string styleName, out int level)
    {
        level = 0;
        if (styleName.Length == 2 &&
            (styleName[0] == 'h' || styleName[0] == 'H') &&
            char.IsDigit(styleName[1]))
        {
            var parsed = styleName[1] - '0';
            if (parsed is >= 1 and <= 6)
            {
                level = parsed;
                return true;
            }
        }
        return false;
    }

    /// <summary>Converts a OneNote run's inline HTML (bold/italic spans, links, breaks) into Markdown.</summary>
    /// <param name="html">The inline HTML fragment from a OneNote text run.</param>
    /// <returns>The equivalent inline Markdown.</returns>
    public string InlineHtmlToMarkdown(string html)
    {
        if (string.IsNullOrWhiteSpace(html)) return string.Empty;
        var doc = new HtmlDocument();
        doc.LoadHtml(html);
        var sb = new StringBuilder();
        ConvertInlineNode(doc.DocumentNode, sb);
        return sb.ToString();
    }

    private void ConvertInlineNode(HtmlNode node, StringBuilder sb)
    {
        foreach (var child in node.ChildNodes)
        {
            if (child.NodeType == HtmlNodeType.Text)
            {
                sb.Append(WebUtility.HtmlDecode(child.InnerText));
                continue;
            }

            if (child.NodeType != HtmlNodeType.Element) continue;

            var name = child.Name.ToLowerInvariant();
            if (name == "br")
            {
                sb.Append('\n');
                continue;
            }

            if (name == "a")
            {
                var linkText = new StringBuilder();
                ConvertInlineNode(child, linkText);
                var href = child.GetAttributeValue("href", string.Empty);
                if (!string.IsNullOrEmpty(href))
                {
                    // Rewrite links that point at another page within this export to a relative
                    // Markdown path; leave all other links (external, cross-notebook) untouched.
                    var rewritten = _linkResolver?.Invoke(WebUtility.HtmlDecode(href));
                    sb.Append('[').Append(linkText).Append("](").Append(rewritten ?? href).Append(')');
                }
                else
                {
                    sb.Append(linkText);
                }
                continue;
            }

            var style = child.GetAttributeValue("style", string.Empty).Replace(" ", string.Empty).ToLowerInvariant();
            var bold = name is "b" or "strong" || style.Contains("font-weight:bold") || style.Contains("font-weight:700");
            var italic = name is "i" or "em" || style.Contains("font-style:italic");
            var strike = name is "s" or "strike" or "del" || style.Contains("text-decoration:line-through");
            var highlight = _highlight != HighlightStyle.None && IsHighlighted(style);

            var inner = new StringBuilder();
            ConvertInlineNode(child, inner);
            var content = inner.ToString();

            // Don't wrap whitespace-only runs; that produces invalid emphasis markers.
            if (!string.IsNullOrWhiteSpace(content))
            {
                if (italic) content = $"*{content}*";
                if (bold) content = $"**{content}**";
                if (strike) content = $"~~{content}~~";
                if (highlight) content = WrapHighlight(content, _highlight);
            }
            sb.Append(content);
        }
    }

    // Detects a OneNote text-highlight, which is encoded as a (non-white) background color on the run.
    private static bool IsHighlighted(string style)
    {
        var value = ExtractCssValue(style, "background-color") ?? ExtractCssValue(style, "background");
        if (string.IsNullOrEmpty(value)) return false;

        // OneNote emits 'background:white' on ordinary (non-highlighted) runs; treat those, and other
        // effectively-absent colors, as no highlight.
        return value is not ("white" or "#ffffff" or "#fff" or "transparent" or "none" or "automatic" or "windowtext");
    }

    // Reads a single CSS declaration value from a space-stripped, lowercased inline style string.
    private static string? ExtractCssValue(string style, string property)
    {
        var key = property + ":";
        var idx = style.IndexOf(key, StringComparison.Ordinal);
        if (idx < 0) return null;

        var start = idx + key.Length;
        var end = style.IndexOf(';', start);
        return end < 0 ? style[start..] : style[start..end];
    }

    private static string WrapHighlight(string content, HighlightStyle style) => style switch
    {
        HighlightStyle.Mark => $"<mark>{content}</mark>",
        HighlightStyle.Equal => $"=={content}==",
        _ => content,
    };

    /// <summary>Decodes OneNote's HTML-encoded run text into trimmed plain text.</summary>
    /// <param name="html">The HTML-encoded run text.</param>
    /// <returns>The decoded plain text.</returns>
    public static string HtmlToPlainText(string html)
    {
        if (string.IsNullOrWhiteSpace(html)) return string.Empty;
        var doc = new HtmlDocument();
        doc.LoadHtml(html);
        return WebUtility.HtmlDecode(doc.DocumentNode.InnerText);
    }
}

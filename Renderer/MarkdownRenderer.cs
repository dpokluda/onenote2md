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

        // Tracks the current ordered-list marker token at each nesting depth (null for bullet levels),
        // used to build hierarchical numbers like "1.1" or "1.a" for the flattened outline.
        public List<string?> NumberStack { get; } = new();
    }

    private void RenderOeChildren(XElement oeChildren, RenderState state, StringBuilder sb, int depth)
    {
        var items = oeChildren.Elements(One + "OE").ToList();
        for (int i = 0; i < items.Count; i++)
        {
            var oe = items[i];

            // Group a run of code-like paragraphs (monospace font, KQL pipe lines, or a JSON object)
            // into a single fenced block. Advances i to the last line consumed.
            if (TryMatchCodeBlock(items, state.Ctx.CodeLanguage, ref i, out var codeLines, out var lang))
            {
                // A KQL query's opening table/operator line is often glued to the end of the preceding
                // prose paragraph (via a soft line break) rather than being its own OE, so a pipe-anchored
                // block loses its first line. If so, reclaim that trailing line from the emitted output.
                if (codeLines.Count > 0 && codeLines[0].TrimStart().StartsWith("|") &&
                    TryReclaimKqlHead(sb, out var head))
                {
                    codeLines.Insert(0, head);
                }

                EnsureBlankLine(sb);
                sb.Append("```").Append(lang).Append('\n');
                sb.Append(string.Join("\n", codeLines).TrimEnd()).Append("\n```\n\n");
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
                RenderListItem(list, text, state, sb, depth);
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

    // Visual indentation added per outline-nesting level. Non-breaking spaces render as real indentation
    // in HTML while never reaching the 4-space threshold that would turn a line into an indented code block.
    private const string ListIndentUnit = "&nbsp;&nbsp;";

    // Renders one outline item as a plain, flattened paragraph line with a computed literal marker,
    // rather than a real Markdown list item. This keeps nesting and (crucially) intra-item code blocks
    // and images robust: each item is an ordinary line, so a code fence between items renders normally
    // and can never be swallowed by, or restart, a Markdown list.
    //   - Bullets:  depth 1 => "*", depth 2 => "**", depth 3 => "***"  (emitted escaped as \* so Markdown
    //               shows the asterisks literally instead of creating a bullet or bold text).
    //   - Numbered: a hierarchical path built from OneNote's per-level markers, e.g. "1", "1.1", "1.a".
    private void RenderListItem(XElement list, string text, RenderState state, StringBuilder sb, int depth)
    {
        var indent = string.Concat(Enumerable.Repeat(ListIndentUnit, depth));
        var body = ToInlineBreaks(text);
        var number = list.Element(One + "Number");

        string marker;
        if (number is null)
        {
            // Bullet level: record a null token so nested numbered items skip this level in their path.
            SetListMarker(state.NumberStack, depth, null);
            marker = string.Concat(Enumerable.Repeat("\\*", depth + 1));
        }
        else
        {
            var token = StripMarkerPunctuation((string?)number.Attribute("text"));
            SetListMarker(state.NumberStack, depth, token);
            marker = string.Join(".", state.NumberStack.Take(depth + 1).Where(t => !string.IsNullOrEmpty(t)));
            if (marker.Length == 0) marker = token;
        }

        sb.Append(indent).Append(marker).Append(' ').Append(body).Append("  \n");
    }

    // Records the marker 'token' (or null for a bullet) at the given nesting depth, discarding any deeper
    // levels so a later item at a shallower depth rebuilds its number path correctly.
    private static void SetListMarker(List<string?> stack, int depth, string? token)
    {
        if (stack.Count > depth + 1)
        {
            stack.RemoveRange(depth + 1, stack.Count - (depth + 1));
        }
        while (stack.Count <= depth)
        {
            stack.Add(null);
        }
        stack[depth] = token;
    }

    // Strips a single trailing '.' or ')' from a OneNote list marker (e.g. "1." => "1", "a)" => "a").
    private static string StripMarkerPunctuation(string? marker)
    {
        var m = (marker ?? string.Empty).Trim();
        if (m.Length > 0 && (m[^1] is '.' or ')'))
        {
            m = m[..^1];
        }
        return m;
    }

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

    // A quoted JSON key (`"name":`), the signal that distinguishes real JSON from bracketed prose.
    private static readonly System.Text.RegularExpressions.Regex JsonKey =
        new("\"[^\"\\r\\n]*\"\\s*:", System.Text.RegularExpressions.RegexOptions.Compiled);

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

    // Characters that break a Markdown inline-link destination when left bare. Wrapping the path in
    // angle brackets (<...>) lets CommonMark parse spaces and (), [] literally.
    private static readonly char[] LinkUnsafeChars = { ' ', '(', ')', '[', ']' };

    // Wraps a relative asset path in angle brackets when it contains characters that would otherwise
    // break the Markdown link so the destination parses correctly.
    private static string ToAssetLink(string relativePath) =>
        relativePath.IndexOfAny(LinkUnsafeChars) >= 0 ? $"<{relativePath}>" : relativePath;

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

    // OneNote records no semantic "code" marker; the only reliable native signal is a monospace font.
    // In practice authors apply it inconsistently, so on top of the font signal we also recognize KQL
    // pipe lines and JSON objects by content, and group adjacent code-like lines into one fenced block.

    // Attempts to consume a fenced code block starting at items[i]. On success, sets 'lines' to the block
    // body, 'lang' to the fence language, advances 'i' to the last consumed item, and returns true.
    private bool TryMatchCodeBlock(List<XElement> items, string defaultLang, ref int i,
        out List<string> lines, out string lang)
    {
        lines = new List<string>();
        lang = defaultLang;

        int start = i;
        if (!IsTextParagraph(items[start])) return false;

        var startTrim = ParagraphText(items[start]).TrimStart();
        var mono = IsMonospaceParagraph(items[start]);
        var pipe = startTrim.StartsWith("|");
        var jsonOpen = startTrim.Length > 0 && (startTrim[0] is '{' or '[');

        // A KQL query's first line (a table name or operator like `union X`) has no leading pipe, so it
        // would otherwise leak out above the block. Treat it as an anchor when the next line is a pipe.
        var nextPipe = start + 1 < items.Count && IsTextParagraph(items[start + 1]) &&
                       ParagraphText(items[start + 1]).TrimStart().StartsWith("|");
        var kqlHead = nextPipe && IsKqlHead(startTrim);

        // A block may only START on a strong anchor: monospace text, a KQL pipe/head line, or a JSON opener.
        if (!(mono || pipe || jsonOpen || kqlHead)) return false;

        var collected = new List<string>();
        int j = start;

        if (jsonOpen && !mono && !pipe)
        {
            // Plain-font JSON: consume until the braces balance. Require a quoted key (`"name":`) rather
            // than any colon so bracketed prose like `[ARM only]:` or `Entity:Diagnostics` isn't fenced.
            int balance = 0;
            for (; j < items.Count; j++)
            {
                if (!IsTextParagraph(items[j])) break;
                var txt = ParagraphText(items[j]);
                if (j > start && !(IsMonospaceParagraph(items[j]) || IsJsonLine(txt.TrimStart()))) break;

                collected.Add(txt);
                foreach (var c in txt)
                {
                    if (c is '{' or '[') balance++;
                    else if (c is '}' or ']') balance--;
                }
                if (balance <= 0) { j++; break; }
            }
            if (collected.Count == 0 || !JsonKey.IsMatch(string.Join("\n", collected))) return false;
        }
        else
        {
            // Monospace / KQL run: the anchor line (start) is always taken; then keep consuming adjacent
            // code-like lines (monospace, pipe, or JSON-ish).
            for (; j < items.Count; j++)
            {
                if (!IsTextParagraph(items[j])) break;
                var txt = ParagraphText(items[j]);
                var trim = txt.TrimStart();
                if (j > start && !(IsMonospaceParagraph(items[j]) || trim.StartsWith("|") || IsJsonLine(trim))) break;
                collected.Add(txt);
            }
            if (collected.Count == 0) return false;
        }

        var firstLine = collected.FirstOrDefault(l => !string.IsNullOrWhiteSpace(l))?.TrimStart() ?? string.Empty;
        if ((firstLine.StartsWith("{") || firstLine.StartsWith("[")) && JsonKey.IsMatch(string.Join("\n", collected)))
        {
            lang = "json";
        }

        lines = collected;
        i = j - 1;
        return true;
    }

    // True when the paragraph is a plain text block (no list marker, table, image, or nested children)
    // that carries at least some non-whitespace text.
    private static bool IsTextParagraph(XElement oe)
    {
        if (oe.Element(One + "List") is not null) return false;
        if (oe.Element(One + "Table") is not null) return false;
        if (oe.Element(One + "Image") is not null) return false;
        if (oe.Elements(One + "OEChildren").Any()) return false;

        var runs = oe.Elements(One + "T").ToList();
        if (runs.Count == 0) return false;
        return runs.Any(t => !string.IsNullOrWhiteSpace(HtmlToPlainText(t.Value)));
    }

    // The paragraph's full plain text, concatenating its inline runs (a single logical line).
    private static string ParagraphText(XElement oe) =>
        string.Concat(oe.Elements(One + "T").Select(t => HtmlToPlainText(t.Value)));

    // A paragraph is monospace when its own style or any run declares a known monospace font-family.
    private static bool IsMonospaceParagraph(XElement oe)
    {
        if (!IsTextParagraph(oe)) return false;

        var runs = oe.Elements(One + "T").ToList();
        var styleHaystack = ((string?)oe.Attribute("style") ?? string.Empty) + " " +
                            string.Concat(runs.Select(t => t.Value));
        styleHaystack = styleHaystack.ToLowerInvariant();
        return MonospaceFonts.Any(f => styleHaystack.Contains("font-family:" + f) || styleHaystack.Contains("font-family: " + f));
    }

    // A JSON-ish continuation line: starts with a bracket/brace or a quoted key/value.
    private static bool IsJsonLine(string trimmed) =>
        trimmed.Length > 0 && (trimmed[0] is '{' or '}' or '[' or ']' or '"');

    private static readonly string[] KqlLeadKeywords =
        { "let ", "union ", "search ", "print ", "range ", "datatable", "find ", "externaldata", "declare ", "set " };

    // A KQL query's opening line: a leading tabular operator/keyword, or a bare (dotted) table identifier.
    // Only used as a code anchor when the following line is a pipe, so the whitelist can stay generous.
    private static bool IsKqlHead(string trimmed)
    {
        if (trimmed.Length == 0) return false;
        var lower = trimmed.ToLowerInvariant();
        if (KqlLeadKeywords.Any(k => lower.StartsWith(k))) return true;
        return trimmed.All(c => char.IsLetterOrDigit(c) || c is '_' or '.');
    }

    // Reclaims a KQL query's opening table/operator line when it was emitted as the trailing line of the
    // preceding prose paragraph (OneNote often glues the table name to the paragraph above the pipes via a
    // soft line break). If the last emitted content line is a KQL head, it is removed from 'sb' and returned
    // so the caller can make it the first line of the following code block.
    private static bool TryReclaimKqlHead(StringBuilder sb, out string head)
    {
        head = string.Empty;
        var text = sb.ToString();

        // Find the end of the last non-whitespace content.
        int end = text.Length;
        while (end > 0 && (text[end - 1] is '\n' or '\r' or ' ' or '\t')) end--;
        if (end == 0) return false;

        int lineStart = text.LastIndexOf('\n', end - 1);
        var candidate = text.Substring(lineStart + 1, end - (lineStart + 1)).Trim();

        // Keep it conservative: a single short token/keyword line, never a hard-broken emphasis/list line.
        if (candidate.Length == 0 || candidate.Length > 60) return false;
        if (candidate.Contains('*') || candidate.Contains('[') || candidate.Contains('`')) return false;
        if (!IsKqlHead(candidate)) return false;

        // Drop the reclaimed line (and any trailing blank/hard-break lines) and re-terminate the paragraph.
        int cut = lineStart < 0 ? 0 : lineStart;
        while (cut > 0 && (sb[cut - 1] is '\n' or '\r' or ' ' or '\t')) cut--;
        sb.Length = cut;
        if (sb.Length > 0) sb.Append("\n\n");

        head = candidate;
        return true;
    }

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
                // Markdown emphasis is inactive when a marker is adjacent to whitespace (e.g. "**ABC **"),
                // so keep any leading/trailing whitespace outside the markers.
                int s = 0, e = content.Length;
                while (s < e && char.IsWhiteSpace(content[s])) s++;
                while (e > s && char.IsWhiteSpace(content[e - 1])) e--;
                var lead = content.Substring(0, s);
                var core = content.Substring(s, e - s);
                var trail = content.Substring(e);

                if (italic) core = $"*{core}*";
                if (bold) core = $"**{core}**";
                if (strike) core = $"~~{core}~~";
                if (highlight) core = WrapHighlight(core, _highlight);

                content = lead + core + trail;
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

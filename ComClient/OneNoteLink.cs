using System.Net;

namespace OneNote2Md.ComClient;

/// <summary>
/// Parses OneNote "Copy Link" URLs to identify the target section's backing <c>.one</c> file and,
/// when the link points at a page, the page's name.
/// </summary>
public static class OneNoteLink
{
    /// <summary>
    /// Extracts a section (and optional page) reference from a OneNote share link. Supports the
    /// <c>onenote:</c> protocol link (which embeds the full <c>.one</c> file URL) and the SharePoint
    /// <c>Doc.aspx?...&amp;wd=target(...)</c> web link (which embeds a notebook-relative path). The
    /// clipboard typically holds both forms; either (or both pasted together) is accepted.
    /// </summary>
    /// <param name="link">The raw link text copied from OneNote.</param>
    /// <returns>The parsed target, or <c>null</c> when no <c>.one</c> reference can be found.</returns>
    public static OneNoteLinkTarget? Parse(string link)
    {
        if (string.IsNullOrWhiteSpace(link)) return null;

        var (fullUrl, protocolPage) = ParseProtocolForm(link);
        var (relativePath, webPage) = ParseWebForm(link);
        var pageId = ExtractPageIdGuid(link);

        // A OneNote page link is usable if it carries the section's .one reference OR a page-id GUID.
        // Intra-notebook links often supply only "#PageName&section-id=...&page-id=..." with no .one
        // URL at all, so the page-id GUID is what makes those resolvable.
        if (fullUrl is null && relativePath is null && pageId is null) return null;

        // Prefer the page name from whichever form supplied one.
        var pageName = protocolPage ?? webPage;
        return new OneNoteLinkTarget(fullUrl, relativePath, pageName, pageId);
    }

    /// <summary>
    /// Extracts the OneNote <c>page-id</c> GUID from a link (or any string containing a
    /// <c>page-id={...}</c> token), normalized to lowercase hex without braces. Handles both the
    /// <c>onenote:</c> form (<c>&amp;page-id={GUID}&amp;</c>) and the SharePoint web form, whose
    /// <c>wd=target(...)</c> value ends with the same page GUID as its final <c>|</c>-delimited segment.
    /// Returns <c>null</c> when no page GUID can be found.
    /// </summary>
    public static string? ExtractPageIdGuid(string? link)
    {
        if (string.IsNullOrEmpty(link)) return null;

        var decoded = WebUtility.UrlDecode(link);

        var idx = decoded.IndexOf("page-id=", StringComparison.OrdinalIgnoreCase);
        if (idx >= 0)
        {
            var rest = decoded[(idx + "page-id=".Length)..];
            var end = FirstIndexOfAny(rest, '&', '}', ')', '\r', '\n', '\t', ' ');
            // Include the closing brace in the captured span when the value is brace-wrapped.
            if (end < rest.Length && rest[end] == '}') end++;
            return NormalizeGuid(rest[..end]);
        }

        return null;
    }

    // Reduces a possibly brace-wrapped GUID string to lowercase hex with no braces or surrounding
    // whitespace, so page GUIDs from different link forms compare equal.
    private static string? NormalizeGuid(string? value)
    {
        value = value?.Trim().Trim('{', '}').Trim();
        return string.IsNullOrEmpty(value) ? null : value.ToLowerInvariant();
    }

    // Parses "onenote:https://.../Section.one#PageName&section-id=...&page-id=...". The .one URL is
    // either inline (right after "onenote:") or, for some links, supplied via a trailing
    // "base-path=" query parameter (e.g. "onenote:#PageName&...&base-path=https://.../Section.one").
    private static (string? OneFileUrl, string? PageName) ParseProtocolForm(string link)
    {
        var idx = link.IndexOf("onenote:", StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return (null, null);

        var rest = link[(idx + "onenote:".Length)..];

        // The inline .one URL (when present) ends at the page fragment ('#'), the first query
        // parameter ('&'), or any whitespace that follows it on the clipboard.
        var urlEnd = FirstIndexOfAny(rest, '#', '&', '\r', '\n', '\t', ' ');
        var inline = WebUtility.UrlDecode(rest[..urlEnd]).Trim();

        // A '#PageName' fragment (present when the link targets a page) runs up to the next '&'.
        string? pageName = null;
        if (urlEnd < rest.Length && rest[urlEnd] == '#')
        {
            var fragment = rest[(urlEnd + 1)..];
            var fragEnd = FirstIndexOfAny(fragment, '&', '\r', '\n', '\t');
            pageName = NormalizePageName(WebUtility.UrlDecode(fragment[..fragEnd]));
        }

        // Keep the parsed page name even when no .one URL is present: intra-notebook links carry a
        // "#PageName" fragment with only section-id/page-id and no .one URL, and the page name is still
        // useful for progress reporting and as a fallback.
        var url = inline.EndsWith(".one", StringComparison.OrdinalIgnoreCase) ? inline : ExtractBasePath(rest);
        return (url, pageName);
    }

    // Extracts the .one URL from a "base-path=https://.../Section.one" query parameter.
    private static string? ExtractBasePath(string rest)
    {
        var idx = rest.IndexOf("base-path=", StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return null;

        var value = rest[(idx + "base-path=".Length)..];
        var end = FirstIndexOfAny(value, '&', '\r', '\n', '\t', ' ');
        var url = WebUtility.UrlDecode(value[..end]).Trim();
        return url.EndsWith(".one", StringComparison.OrdinalIgnoreCase) ? url : null;
    }

    // Parses "...&wd=target(Policy/On-Call.one|<section-guid>/PageName|<page-guid>/)".
    private static (string? RelativeOnePath, string? PageName) ParseWebForm(string link)
    {
        var idx = link.IndexOf("wd=", StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return (null, null);

        var rest = link[(idx + "wd=".Length)..];
        var amp = rest.IndexOf('&');
        var value = WebUtility.UrlDecode(amp >= 0 ? rest[..amp] : rest);

        var open = value.IndexOf('(');
        if (open < 0) return (null, null);
        var close = value.IndexOf(')', open + 1);
        var inner = close > open ? value[(open + 1)..close] : value[(open + 1)..];

        // inner = "<relpath>.one|<section-guid>/PageName|<page-guid>/"
        var segments = inner.Split('|');

        var path = segments[0].Replace('\\', '/').Trim().Trim('/');
        if (!path.EndsWith(".one", StringComparison.OrdinalIgnoreCase)) return (null, null);

        // The second segment is "<section-guid>/PageName"; everything after the first '/' is the page.
        string? pageName = null;
        if (segments.Length > 1)
        {
            var slash = segments[1].IndexOf('/');
            if (slash >= 0) pageName = NormalizePageName(segments[1][(slash + 1)..]);
        }

        return (path, pageName);
    }

    private static int FirstIndexOfAny(string value, params char[] chars)
    {
        var end = value.Length;
        foreach (var ch in chars)
        {
            var at = value.IndexOf(ch);
            if (at >= 0 && at < end) end = at;
        }
        return end;
    }

    private static string? NormalizePageName(string? name)
    {
        name = name?.Trim();
        return string.IsNullOrEmpty(name) ? null : name;
    }
}

/// <summary>Identifies the section file (and optional page) referenced by a OneNote link.</summary>
/// <param name="OneFileUrl">Full <c>.one</c> file URL from an <c>onenote:</c> link, or <c>null</c>.</param>
/// <param name="RelativeOnePath">Notebook-relative <c>.one</c> path from a web link, or <c>null</c>.</param>
/// <param name="PageName">Target page name when the link points at a page, or <c>null</c> for a section link.</param>
/// <param name="PageId">Normalized <c>page-id</c> GUID (lowercase, no braces) when the link targets a
/// page, or <c>null</c>. This is the most reliable page identifier, since it survives page renames
/// and is independent of the section's <c>.one</c> URL.</param>
public sealed record OneNoteLinkTarget(
    string? OneFileUrl,
    string? RelativeOnePath,
    string? PageName,
    string? PageId = null);

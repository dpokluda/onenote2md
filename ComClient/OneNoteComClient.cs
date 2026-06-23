using System.Globalization;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;

namespace OneNote2Md.ComClient;

/// <summary>
/// Reads the OneNote hierarchy and page content from the local OneNote desktop client via COM
/// automation (PIA-bound). Requires OneNote 2016+ installed and the target notebook already
/// opened on this machine. Sections and section groups are identified by their hierarchical
/// display path (notebook/group/.../name) because OneNote assigns its own internal IDs.
/// </summary>
public sealed class OneNoteComClient
{
    internal static readonly XNamespace One = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    // The OneNote COM Application is single-threaded; serialize every call through this lock.
    private readonly object _comLock = new();
    private Application? _app;

    // Lazily connects to the running OneNote desktop instance, reusing the COM object thereafter.
    private Application App
    {
        get
        {
            if (_app is not null) return _app;
            lock (_comLock)
            {
                if (_app is not null) return _app;
                _app = ConnectWithRetry();
                return _app;
            }
        }
    }

    // OneNote can be slow to cold-launch its COM server, so retry a few times before giving up.
    private static Application ConnectWithRetry()
    {
        const int attempts = 3;
        COMException? last = null;

        for (var attempt = 1; attempt <= attempts; attempt++)
        {
            try
            {
                return new Application();
            }
            catch (COMException ex)
            {
                last = ex;
                if (attempt < attempts) Thread.Sleep(1500);
            }
        }

        throw new InvalidOperationException(DescribeConnectionFailure(last!), last);
    }

    // Turns a raw COM HRESULT into actionable guidance for the most common OneNote failures.
    private static string DescribeConnectionFailure(COMException ex)
    {
        var hint = unchecked((uint)ex.HResult) switch
        {
            0x80080005 => // CO_E_SERVER_EXEC_FAILURE
                "Windows could not start the OneNote COM server. The usual cause is an elevation " +
                "mismatch: run this tool from a normal (non-Administrator) terminal when OneNote is " +
                "running normally (or run both elevated). Also make sure the desktop OneNote " +
                "(Microsoft 365 / 2016) is installed and running \u2014 the 'OneNote for Windows 10' " +
                "Store app does not support automation.",
            0x800401E3 => // MK_E_UNAVAILABLE
                "OneNote is not running. Start the desktop OneNote (Microsoft 365 / 2016) and open " +
                "the target notebook, then try again.",
            0x80040154 => // REGDB_E_CLASSNOTREG
                "The OneNote automation component is not registered. Install or repair the desktop " +
                "OneNote (Microsoft 365 / 2016); the 'OneNote for Windows 10' Store app is not supported.",
            _ => "Make sure the desktop OneNote (Microsoft 365 / 2016) is installed and running with " +
                 "the target notebook open, and that this tool runs at the same elevation level as OneNote.",
        };

        return $"Could not connect to OneNote desktop (0x{(uint)ex.HResult:X8}). {hint}";
    }

    /// <summary>
    /// Eagerly establishes the COM connection to OneNote so connection failures surface up front
    /// rather than on the first hierarchy call.
    /// </summary>
    public void Connect()
    {
        _ = App;
    }

    /// <summary>
    /// Resolves the supplied display path to a section or section group and builds the export
    /// subtree rooted at it. Returns <c>null</c> when the path cannot be found.
    /// </summary>
    /// <param name="path">Hierarchical OneNote path, e.g. "Notebook/Group/Section".</param>
    /// <returns>The root node of the resolved subtree, or <c>null</c> if the path does not exist.</returns>
    public OneNoteNode? ResolveTarget(string path)
    {
        var hierarchy = LoadHierarchy();
        var element = ResolveElementByPath(hierarchy, path);
        if (element is null) return null;

        var name = (string?)element.Attribute("name") ?? path;
        return element.Name == One + "Section"
            ? BuildSection(element)
            : BuildGroup(element, name);
    }

    /// <summary>
    /// Resolves the section or page referenced by a OneNote "Copy Link" URL by matching the link's
    /// backing <c>.one</c> file against the OneNote hierarchy. When the link points at a page, the
    /// page (and its sub-pages) is returned; otherwise the whole section is returned.
    /// </summary>
    /// <param name="link">The raw OneNote share link (the <c>onenote:</c> protocol and/or SharePoint web form).</param>
    /// <param name="sourcePath">On success, the hierarchical display path of the export root's container (the section).</param>
    /// <param name="failure">On failure, a human-readable explanation; otherwise <c>null</c>.</param>
    /// <returns>The resolved section or page node, or <c>null</c> if the link cannot be parsed or matched.</returns>
    public OneNoteNode? ResolveTargetByLink(string link, out string sourcePath, out string? failure)
    {
        sourcePath = string.Empty;
        failure = null;

        var target = OneNoteLink.Parse(link);
        if (target is null)
        {
            failure = "The link does not contain a recognizable OneNote section reference.";
            return null;
        }

        var hierarchy = LoadHierarchy();
        var sectionElement = FindSectionByOneFile(hierarchy, target);
        if (sectionElement is null)
        {
            failure = "Could not match the link's section file to an open notebook. " +
                      "Make sure the notebook is open in OneNote.";
            return null;
        }

        var sectionPath = BuildDisplayPath(sectionElement);
        var section = BuildSection(sectionElement);
        sourcePath = sectionPath;

        if (target.PageName is null)
        {
            return section;
        }

        var page = FindPageByName(section.Pages, target.PageName);
        if (page is null)
        {
            failure = $"Found the section '{sectionPath}' but no page named '{target.PageName}'. " +
                      "The page may have been renamed or moved to a different section.";
            return null;
        }

        return page;
    }

    // Searches a page tree by name, preferring shallower matches (top-level pages before sub-pages).
    private static PageNode? FindPageByName(IReadOnlyList<PageNode> pages, string name)
    {
        foreach (var page in pages)
        {
            if (string.Equals(page.Name, name, StringComparison.OrdinalIgnoreCase)) return page;
        }

        foreach (var page in pages)
        {
            var found = FindPageByName(page.SubPages, name);
            if (found is not null) return found;
        }

        return null;
    }

    /// <summary>Returns the raw OneNote XML for a single page, including inlined binary image data.</summary>
    /// <param name="pageId">OneNote internal page ID to fetch.</param>
    /// <returns>The page content as OneNote XML.</returns>
    public string GetPageXml(string pageId)
    {
        lock (_comLock)
        {
            App.GetPageContent(pageId, out var xml, PageInfo.piBinaryData);
            return xml;
        }
    }

    // ---- hierarchy loading -------------------------------------------------

    // Fetches the page-level hierarchy (notebooks > groups > sections > pages) as XML.
    private XDocument LoadHierarchy()
    {
        string xml;
        lock (_comLock)
        {
            App.GetHierarchy(null, HierarchyScope.hsPages, out xml);
        }

        if (string.IsNullOrEmpty(xml))
        {
            throw new InvalidOperationException(
                "OneNote returned an empty hierarchy. Is OneNote desktop running with notebooks loaded?");
        }

        return XDocument.Parse(xml);
    }

    // Walks the hierarchy XML from a notebook down through section groups to the final segment,
    // which may be either a Section or a SectionGroup.
    private static XElement? ResolveElementByPath(XDocument hierarchy, string path)
    {
        var parts = path.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length < 1) return null;

        var notebook = hierarchy.Descendants(One + "Notebook")
            .FirstOrDefault(n => NameEquals(n, parts[0]));
        if (notebook is null) return null;

        XElement current = notebook;
        for (int i = 1; i < parts.Length; i++)
        {
            var isLast = i == parts.Length - 1;
            var group = current.Elements(One + "SectionGroup").FirstOrDefault(g => NameEquals(g, parts[i]));

            if (isLast)
            {
                // The final segment is either a section or a section group.
                var section = current.Elements(One + "Section").FirstOrDefault(s => NameEquals(s, parts[i]));
                return section ?? group;
            }

            if (group is null) return null;
            current = group;
        }

        // Path was just the notebook name: treat the notebook as a section group.
        return notebook;
    }

    private static bool NameEquals(XElement element, string name) =>
        string.Equals((string?)element.Attribute("name"), name, StringComparison.OrdinalIgnoreCase);

    // Finds the Section whose backing .one file matches the link target. An exact full-URL match
    // wins outright; a notebook-relative path match is used only as a fallback.
    private static XElement? FindSectionByOneFile(XDocument hierarchy, OneNoteLinkTarget target)
    {
        XElement? relativeMatch = null;

        foreach (var section in hierarchy.Descendants(One + "Section"))
        {
            if ((bool?)section.Attribute("isInRecycleBin") == true) continue;

            var path = (string?)section.Attribute("path");
            if (string.IsNullOrEmpty(path)) continue;

            var normalized = path.Replace('\\', '/');

            if (target.OneFileUrl is not null &&
                string.Equals(normalized, target.OneFileUrl, StringComparison.OrdinalIgnoreCase))
            {
                return section;
            }

            if (relativeMatch is null && target.RelativeOnePath is not null &&
                normalized.EndsWith("/" + target.RelativeOnePath, StringComparison.OrdinalIgnoreCase))
            {
                relativeMatch = section;
            }
        }

        return relativeMatch;
    }

    // Reconstructs a Notebook/Group/.../Section display path by walking the element's ancestors.
    private static string BuildDisplayPath(XElement section)
    {
        var names = section.AncestorsAndSelf()
            .Where(e => e.Name == One + "Notebook" || e.Name == One + "SectionGroup" || e.Name == One + "Section")
            .Select(e => (string?)e.Attribute("name"))
            .Where(n => !string.IsNullOrEmpty(n))
            .Reverse();
        return string.Join("/", names);
    }

    // ---- model building ----------------------------------------------------

    private static SectionGroupNode BuildGroup(XElement element, string name)
    {
        var node = new SectionGroupNode { Name = name };

        foreach (var group in element.Elements(One + "SectionGroup"))
        {
            // OneNote stores the recycle bin as a hidden section group; skip it.
            if ((bool?)group.Attribute("isRecycleBin") == true) continue;
            node.Groups.Add(BuildGroup(group, (string?)group.Attribute("name") ?? "(unnamed group)"));
        }

        foreach (var section in element.Elements(One + "Section"))
        {
            node.Sections.Add(BuildSection(section));
        }

        return node;
    }

    private static SectionNode BuildSection(XElement element)
    {
        var node = new SectionNode { Name = (string?)element.Attribute("name") ?? "(unnamed section)" };

        // The .one file URL identifies this section when matching intra-export page links.
        var oneFilePath = ((string?)element.Attribute("path"))?.Replace('\\', '/');

        // Pages arrive as a flat list ordered by appearance; pageLevel encodes sub-page nesting.
        var stack = new Stack<PageNode>();
        foreach (var pageElement in element.Elements(One + "Page"))
        {
            var level = (int?)pageElement.Attribute("pageLevel") ?? 1;
            var page = new PageNode
            {
                Id = (string)pageElement.Attribute("ID")!,
                Name = (string?)pageElement.Attribute("name") ?? "(untitled)",
                OneFilePath = oneFilePath,
                Level = level,
                CreatedUtc = ParseTime(pageElement.Attribute("dateTime")?.Value),
                LastModifiedUtc = ParseTime(pageElement.Attribute("lastModifiedTime")?.Value),
            };

            while (stack.Count > 0 && stack.Peek().Level >= level) stack.Pop();

            if (stack.Count == 0)
            {
                node.Pages.Add(page);
            }
            else
            {
                stack.Peek().SubPages.Add(page);
            }

            stack.Push(page);
        }

        return node;
    }

    private static DateTimeOffset ParseTime(string? value)
    {
        if (string.IsNullOrEmpty(value)) return DateTimeOffset.MinValue;
        return DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var dto)
            ? dto
            : DateTimeOffset.MinValue;
    }
}
